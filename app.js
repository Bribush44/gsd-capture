"use strict";

document.addEventListener("DOMContentLoaded", async function () {
  const TASKS_KEY = "gsdCaptureTasks";
  const SETTINGS_KEY = "gsdCaptureSettings";

  const MICROSOFT_CLIENT_ID =
    "ab018161-147f-41ab-8859-657a46519c37";
  const MICROSOFT_AUTHORITY =
    "https://login.microsoftonline.com/common";
  const MICROSOFT_SCOPES = [
    "User.Read",
    "Tasks.ReadWrite",
  ];
  const MICROSOFT_REVIEW_LIST_NAME = "GSD Review";

  const VOICE_SETUP_PENDING_KEY =
    "gsdVoiceSetupPending";
  const VOICE_SETUP_CONNECTION_KEY =
    "gsdVoiceSetupConnection";

  const captureForm = document.getElementById("captureForm");
  const taskInput = document.getElementById("taskInput");
  const captureMessage = document.getElementById("captureMessage");
  const recentTaskList = document.getElementById("recentTaskList");
  const todayTaskList = document.getElementById("todayTaskList");
  const reviewTaskList = document.getElementById("reviewTaskList");
  const refreshReviewButton = document.getElementById(
    "refreshReviewButton"
  );
  const serverReviewStatus = document.getElementById(
    "serverReviewStatus"
  );

  const inboxCount = document.getElementById("inboxCount");
  const reviewCount = document.getElementById("reviewCount");
  const capturedTotal = document.getElementById("capturedTotal");
  const completedTotal = document.getElementById("completedTotal");
  const needsReviewTotal = document.getElementById(
    "needsReviewTotal"
  );

  const settingsButton = document.getElementById("settingsButton");
  const saveSettingsButton = document.getElementById(
    "saveSettingsButton"
  );

  const userNameInput = document.getElementById("userName");
  const defaultListInput = document.getElementById("defaultList");
  const aiSortingToggle = document.getElementById(
    "aiSortingToggle"
  );
  const calendarToggle = document.getElementById(
    "calendarToggle"
  );
  const settingsMessage = document.getElementById(
    "settingsMessage"
  );

  const connectMicrosoftButton = document.getElementById(
    "connectMicrosoftButton"
  );
  const disconnectMicrosoftButton = document.getElementById(
    "disconnectMicrosoftButton"
  );
  const microsoftStatus = document.getElementById(
    "microsoftStatus"
  );
  const microsoftConnectionBadge = document.getElementById(
    "microsoftConnectionBadge"
  );
  const microsoftListPanel = document.getElementById(
    "microsoftListPanel"
  );
  const microsoftListSelect = document.getElementById(
    "microsoftListSelect"
  );
  const loadMicrosoftListsButton = document.getElementById(
    "loadMicrosoftListsButton"
  );
  const microsoftListMessage = document.getElementById(
    "microsoftListMessage"
  );
  const createMicrosoftTestTaskButton = document.getElementById(
    "createMicrosoftTestTaskButton"
  );
  const microsoftTestTaskMessage = document.getElementById(
    "microsoftTestTaskMessage"
  );
  const microsoftRoutingModeSelect = document.getElementById(
    "microsoftRoutingMode"
  );
  const microsoftRoutingMessage = document.getElementById(
    "microsoftRoutingMessage"
  );

  const voiceSetupCard = document.getElementById(
    "voiceSetupCard"
  );
  const voiceSetupBadge = document.getElementById(
    "voiceSetupBadge"
  );
  const voiceSetupStatus = document.getElementById(
    "voiceSetupStatus"
  );
  const startVoiceSetupButton = document.getElementById(
    "startVoiceSetupButton"
  );
  const voiceSetupResult = document.getElementById(
    "voiceSetupResult"
  );
  const voiceSetupKey = document.getElementById(
    "voiceSetupKey"
  );
  const copyVoiceSetupKeyButton = document.getElementById(
    "copyVoiceSetupKeyButton"
  );

  const voiceButton = document.getElementById("voiceButton");
  const captureButton = captureForm
    ? captureForm.querySelector('button[type="submit"]')
    : null;

  const screens = document.querySelectorAll(".screen");
  const navButtons = document.querySelectorAll(".nav-button");

  let tasks = loadTasks();
  let settings = loadSettings();
  let captureInProgress = false;
  let microsoftAuth = null;
  let microsoftAccount = null;
  let microsoftAuthReady = false;
  let microsoftTodoLists = [];
  let microsoftActionQueueRunning = false;
  let serverReviewLoading = false;

  if (
    !captureForm ||
    !taskInput ||
    !recentTaskList ||
    !todayTaskList ||
    !reviewTaskList
  ) {
    window.alert(
      "The app could not start because part of index.html is missing."
    );
    return;
  }

  populateSettings();
  registerEvents();
  renderEverything();
  await initializeMicrosoftSignIn();
  await resumeVoiceSetup();
  await loadServerReviewInbox();
  await processPendingMicrosoftActions();

  function registerEvents() {
    captureForm.addEventListener("submit", captureTask);

    navButtons.forEach(function (button) {
      button.addEventListener("click", function () {
        showScreen(button.dataset.screen);

        if (button.dataset.screen === "reviewScreen") {
          loadServerReviewInbox();
        }
      });
    });

    if (settingsButton) {
      settingsButton.addEventListener("click", function () {
        showScreen("settingsScreen");

        navButtons.forEach(function (button) {
          button.classList.remove("active-nav");
        });
      });
    }

    if (saveSettingsButton) {
      saveSettingsButton.addEventListener(
        "click",
        saveSettings
      );
    }

    if (voiceButton) {
      voiceButton.addEventListener(
        "click",
        startVoiceCapture
      );
    }

    if (connectMicrosoftButton) {
      connectMicrosoftButton.addEventListener(
        "click",
        connectMicrosoftAccount
      );
    }

    if (disconnectMicrosoftButton) {
      disconnectMicrosoftButton.addEventListener(
        "click",
        disconnectMicrosoftAccount
      );
    }

    if (loadMicrosoftListsButton) {
      loadMicrosoftListsButton.addEventListener(
        "click",
        loadMicrosoftTodoLists
      );
    }

    if (microsoftListSelect) {
      microsoftListSelect.addEventListener(
        "change",
        saveMicrosoftListSelection
      );
    }

    if (createMicrosoftTestTaskButton) {
      createMicrosoftTestTaskButton.addEventListener(
        "click",
        createMicrosoftTestTask
      );
    }

    if (microsoftRoutingModeSelect) {
      microsoftRoutingModeSelect.addEventListener(
        "change",
        saveMicrosoftRoutingMode
      );
    }

    if (startVoiceSetupButton) {
      startVoiceSetupButton.addEventListener(
        "click",
        startVoiceSetup
      );
    }

    if (copyVoiceSetupKeyButton) {
      copyVoiceSetupKeyButton.addEventListener(
        "click",
        copyVoiceSetupKey
      );
    }

    if (refreshReviewButton) {
      refreshReviewButton.addEventListener(
        "click",
        function () {
          loadServerReviewInbox(true);
        }
      );
    }

    window.addEventListener("online", function () {
      loadServerReviewInbox();
      processPendingMicrosoftActions();
    });
  }

  async function initializeMicrosoftSignIn() {
    if (!window.msal || !window.msal.PublicClientApplication) {
      setMicrosoftMessage(
        "Microsoft sign-in could not load. The rest of GSD Capture will still work.",
        true
      );
      return;
    }

    try {
      const redirectUri = window.location.origin + "/";

      microsoftAuth = new window.msal.PublicClientApplication({
        auth: {
          clientId: MICROSOFT_CLIENT_ID,
          authority: MICROSOFT_AUTHORITY,
          redirectUri: redirectUri,
          postLogoutRedirectUri: redirectUri,
          navigateToLoginRequestUrl: false,
        },
        cache: {
          cacheLocation: "localStorage",
        },
      });

      await microsoftAuth.initialize();

      const redirectResult =
        await microsoftAuth.handleRedirectPromise();

      if (redirectResult && redirectResult.account) {
        microsoftAuth.setActiveAccount(
          redirectResult.account
        );
        microsoftAccount = redirectResult.account;
      } else {
        microsoftAccount =
          microsoftAuth.getActiveAccount();

        if (!microsoftAccount) {
          const accounts = microsoftAuth.getAllAccounts();

          if (accounts.length > 0) {
            microsoftAccount = accounts[0];
            microsoftAuth.setActiveAccount(
              microsoftAccount
            );
          }
        }
      }

      microsoftAuthReady = true;
      updateMicrosoftConnectionDisplay();

      if (redirectResult && redirectResult.account) {
        showScreen("settingsScreen");

        navButtons.forEach(function (button) {
          button.classList.remove("active-nav");
        });
      }
    } catch (error) {
      console.error(
        "Microsoft sign-in initialization failed:",
        error
      );

      setMicrosoftMessage(
        "Microsoft sign-in did not initialize. GSD Capture is still available for local capture.",
        true
      );
    }
  }

  async function connectMicrosoftAccount() {
    if (!microsoftAuthReady || !microsoftAuth) {
      setMicrosoftMessage(
        "Microsoft sign-in is still loading. Try again in a moment.",
        true
      );
      return;
    }

    if (connectMicrosoftButton) {
      connectMicrosoftButton.disabled = true;
      connectMicrosoftButton.textContent =
        "Opening Microsoft...";
    }

    setMicrosoftMessage(
      "Opening Microsoft sign-in..."
    );

    try {
      await microsoftAuth.loginRedirect({
        scopes: MICROSOFT_SCOPES,
        prompt: "select_account",
      });
    } catch (error) {
      console.error("Microsoft sign-in failed:", error);

      setMicrosoftMessage(
        "Microsoft sign-in could not start. Please try again.",
        true
      );

      updateMicrosoftConnectionDisplay();
    }
  }

  async function disconnectMicrosoftAccount() {
    if (!microsoftAuthReady || !microsoftAuth) {
      return;
    }

    const account =
      microsoftAccount || microsoftAuth.getActiveAccount();

    setMicrosoftMessage(
      "Disconnecting Microsoft account..."
    );

    try {
      await microsoftAuth.logoutRedirect({
        account: account || undefined,
        postLogoutRedirectUri:
          window.location.origin + "/",
      });
    } catch (error) {
      console.error("Microsoft sign-out failed:", error);

      setMicrosoftMessage(
        "Microsoft sign-out did not complete. Please try again.",
        true
      );
    }
  }

  function updateMicrosoftConnectionDisplay() {
    if (!microsoftAuthReady) {
      return;
    }

    if (microsoftAccount) {
      const displayName =
        microsoftAccount.name ||
        microsoftAccount.username ||
        "Microsoft user";

      if (microsoftConnectionBadge) {
        microsoftConnectionBadge.textContent =
          "Connected";
        microsoftConnectionBadge.classList.add(
          "connected"
        );
      }

      if (connectMicrosoftButton) {
        connectMicrosoftButton.classList.add("hidden");
        connectMicrosoftButton.disabled = false;
        connectMicrosoftButton.textContent =
          "Connect Microsoft To Do";
      }

      if (disconnectMicrosoftButton) {
        disconnectMicrosoftButton.classList.remove(
          "hidden"
        );
      }

      if (microsoftListPanel) {
        microsoftListPanel.classList.remove("hidden");
      }

      if (settings.microsoftListName) {
        setMicrosoftListMessage(
          "Saved destination: " + settings.microsoftListName + ". Load your lists to verify it."
        );
      } else {
        setMicrosoftListMessage(
          "Load your Microsoft To Do lists, then choose the destination list."
        );
      }

      setMicrosoftMessage(
        "Connected as " + displayName + "."
      );

      updateVoiceSetupDisplay();
      return;
    }

    if (microsoftConnectionBadge) {
      microsoftConnectionBadge.textContent =
        "Not connected";
      microsoftConnectionBadge.classList.remove(
        "connected"
      );
    }

    if (connectMicrosoftButton) {
      connectMicrosoftButton.classList.remove("hidden");
      connectMicrosoftButton.disabled = false;
      connectMicrosoftButton.textContent =
        "Connect Microsoft To Do";
    }

    if (disconnectMicrosoftButton) {
      disconnectMicrosoftButton.classList.add("hidden");
    }

    if (microsoftListPanel) {
      microsoftListPanel.classList.add("hidden");
    }

    clearMicrosoftListOptions();

    setMicrosoftMessage(
      "Sign in to connect GSD Capture to your Microsoft account."
    );

    updateVoiceSetupDisplay();
  }

  function loadSavedVoiceSetup() {
    try {
      const saved = localStorage.getItem(
        VOICE_SETUP_CONNECTION_KEY
      );

      if (!saved) {
        return null;
      }

      const parsed = JSON.parse(saved);

      return parsed &&
        typeof parsed === "object" &&
        typeof parsed.voiceKey === "string"
        ? parsed
        : null;
    } catch (error) {
      console.error(
        "Saved Voice Capture setup could not be read:",
        error
      );

      return null;
    }
  }

  function updateVoiceSetupDisplay() {
    if (!voiceSetupCard) {
      return;
    }

    const savedSetup = loadSavedVoiceSetup();

    if (savedSetup && savedSetup.voiceKey) {
      if (voiceSetupBadge) {
        voiceSetupBadge.textContent = "Connected";
        voiceSetupBadge.classList.add("connected");
      }

      if (voiceSetupStatus) {
        const accountName =
          savedSetup.displayName ||
          savedSetup.email ||
          "your Microsoft account";

        voiceSetupStatus.textContent =
          "Voice Capture is connected to " +
          accountName +
          ".";
        voiceSetupStatus.classList.remove(
          "error-message"
        );
      }

      if (startVoiceSetupButton) {
        startVoiceSetupButton.disabled =
          !microsoftAccount;
        startVoiceSetupButton.textContent =
          "Set Up Again";
      }

      if (voiceSetupResult) {
        voiceSetupResult.classList.remove("hidden");
      }

      if (voiceSetupKey) {
        voiceSetupKey.value = savedSetup.voiceKey;
      }

      return;
    }

    if (voiceSetupBadge) {
      voiceSetupBadge.textContent = "Not set up";
      voiceSetupBadge.classList.remove("connected");
    }

    if (voiceSetupResult) {
      voiceSetupResult.classList.add("hidden");
    }

    if (voiceSetupKey) {
      voiceSetupKey.value = "";
    }

    if (startVoiceSetupButton) {
      startVoiceSetupButton.disabled =
        !microsoftAccount;
      startVoiceSetupButton.textContent =
        "Set Up Voice Capture";
    }

    if (voiceSetupStatus) {
      voiceSetupStatus.classList.remove(
        "error-message"
      );

      voiceSetupStatus.textContent =
        microsoftAccount
          ? "Connect private background capture for Siri and Apple Shortcuts."
          : "Connect Microsoft To Do first, then set up private background capture.";
    }
  }

  function setVoiceSetupMessage(message, isError) {
    if (!voiceSetupStatus) {
      return;
    }

    voiceSetupStatus.textContent = message;
    voiceSetupStatus.classList.toggle(
      "error-message",
      Boolean(isError)
    );
  }

  function setVoiceSetupBusy(isBusy) {
    if (!startVoiceSetupButton) {
      return;
    }

    startVoiceSetupButton.disabled =
      Boolean(isBusy) || !microsoftAccount;

    if (isBusy) {
      startVoiceSetupButton.textContent =
        "Starting setup...";
      return;
    }

    startVoiceSetupButton.textContent =
      loadSavedVoiceSetup()
        ? "Set Up Again"
        : "Set Up Voice Capture";
  }

  async function startVoiceSetup() {
    if (!navigator.onLine) {
      setVoiceSetupMessage(
        "Connect to the internet before setting up Voice Capture.",
        true
      );
      return;
    }

    if (
      !microsoftAuthReady ||
      !microsoftAuth ||
      !microsoftAccount
    ) {
      setVoiceSetupMessage(
        "Connect your Microsoft account first.",
        true
      );
      return;
    }

    setVoiceSetupBusy(true);
    setVoiceSetupMessage(
      "Verifying your Microsoft account..."
    );

    try {
      const accessToken =
        await acquireMicrosoftAccessToken();

      if (!accessToken) {
        setVoiceSetupBusy(false);
        return;
      }

      const response = await fetch(
        "/api/voice-setup-start",
        {
          method: "POST",
          headers: {
            Authorization:
              "Bearer " + accessToken,
            "Content-Type":
              "application/json",
            Accept: "application/json",
          },
          body: JSON.stringify({}),
        }
      );

      const data = await response
        .json()
        .catch(function () {
          return null;
        });

      if (
        !response.ok ||
        !data ||
        !data.setupId ||
        !data.voiceKey ||
        !data.connectUrl
      ) {
        throw new Error(
          data &&
            (data.message || data.error)
            ? data.message || data.error
            : "Voice Capture setup could not be started."
        );
      }

      const pendingSetup = {
        setupId: data.setupId,
        voiceKey: data.voiceKey,
        displayName: data.displayName || "",
        email: data.email || "",
        connectUrl: data.connectUrl,
        expiresAt:
          Date.now() +
          Number(data.expiresInSeconds || 900) *
            1000,
      };

      sessionStorage.setItem(
        VOICE_SETUP_PENDING_KEY,
        JSON.stringify(pendingSetup)
      );

      setVoiceSetupMessage(
        "Opening Microsoft to approve background capture..."
      );

      window.location.assign(data.connectUrl);
    } catch (error) {
      console.error(
        "Voice Capture setup failed to start:",
        error
      );

      setVoiceSetupMessage(
        error && error.message
          ? error.message
          : "Voice Capture setup could not be started. Please try again.",
        true
      );

      setVoiceSetupBusy(false);
    }
  }

  async function resumeVoiceSetup() {
    updateVoiceSetupDisplay();

    const parameters = new URLSearchParams(
      window.location.search
    );

    const returnedSetupId =
      parameters.get("setupId") || "";
    const returnedComplete =
      parameters.get("voiceSetup") === "complete";

    let pendingSetup = null;

    try {
      const savedPending = sessionStorage.getItem(
        VOICE_SETUP_PENDING_KEY
      );

      pendingSetup = savedPending
        ? JSON.parse(savedPending)
        : null;
    } catch (error) {
      console.error(
        "Pending Voice Capture setup could not be read:",
        error
      );
    }

    if (!returnedComplete && !pendingSetup) {
      return;
    }

    showScreen("settingsScreen");

    navButtons.forEach(function (button) {
      button.classList.remove("active-nav");
    });

    if (
      !pendingSetup ||
      !pendingSetup.setupId ||
      !pendingSetup.voiceKey
    ) {
      setVoiceSetupMessage(
        "The private Voice Capture key was not found. Start setup again.",
        true
      );

      clearVoiceSetupQuery();
      return;
    }

    if (
      returnedSetupId &&
      returnedSetupId !== pendingSetup.setupId
    ) {
      setVoiceSetupMessage(
        "The returned Voice Capture setup did not match this device. Start setup again.",
        true
      );

      clearVoiceSetupQuery();
      return;
    }

    if (
      pendingSetup.expiresAt &&
      Date.now() > pendingSetup.expiresAt
    ) {
      sessionStorage.removeItem(
        VOICE_SETUP_PENDING_KEY
      );

      setVoiceSetupMessage(
        "Voice Capture setup expired. Start setup again.",
        true
      );

      clearVoiceSetupQuery();
      return;
    }

    setVoiceSetupBusy(true);
    setVoiceSetupMessage(
      "Confirming your Voice Capture connection..."
    );

    try {
      const response = await fetch(
        "/api/voice-setup-status?setup=" +
          encodeURIComponent(
            pendingSetup.setupId
          ),
        {
          headers: {
            Accept: "application/json",
          },
          cache: "no-store",
        }
      );

      const data = await response
        .json()
        .catch(function () {
          return null;
        });

      if (!response.ok || !data) {
        throw new Error(
          data && data.error
            ? data.error
            : "Voice Capture status could not be checked."
        );
      }

      if (!data.complete) {
        if (data.expired) {
          throw new Error(
            "Voice Capture setup expired. Start setup again."
          );
        }

        setVoiceSetupMessage(
          "Microsoft approval has not finished yet. Return here after completing it.",
          true
        );

        setVoiceSetupBusy(false);
        return;
      }

      const completedSetup = {
        voiceKey: pendingSetup.voiceKey,
        setupId: pendingSetup.setupId,
        displayName:
          data.displayName ||
          pendingSetup.displayName ||
          "",
        email:
          data.email ||
          pendingSetup.email ||
          "",
        connectedAt:
          data.connectedAt ||
          new Date().toISOString(),
      };

      localStorage.setItem(
        VOICE_SETUP_CONNECTION_KEY,
        JSON.stringify(completedSetup)
      );

      sessionStorage.removeItem(
        VOICE_SETUP_PENDING_KEY
      );

      clearVoiceSetupQuery();
      updateVoiceSetupDisplay();

      setVoiceSetupMessage(
        "Voice Capture is connected. Copy the private key for the Shortcut installation step."
      );
    } catch (error) {
      console.error(
        "Voice Capture setup confirmation failed:",
        error
      );

      setVoiceSetupMessage(
        error && error.message
          ? error.message
          : "Voice Capture setup could not be confirmed.",
        true
      );
    } finally {
      setVoiceSetupBusy(false);
    }
  }

  function clearVoiceSetupQuery() {
    const url = new URL(window.location.href);

    url.searchParams.delete("voiceSetup");
    url.searchParams.delete("setupId");

    window.history.replaceState(
      {},
      document.title,
      url.pathname +
        url.search +
        url.hash
    );
  }

  async function copyVoiceSetupKey() {
    const savedSetup = loadSavedVoiceSetup();

    if (!savedSetup || !savedSetup.voiceKey) {
      setVoiceSetupMessage(
        "No Voice Capture key is available yet.",
        true
      );
      return;
    }

    try {
      await navigator.clipboard.writeText(
        savedSetup.voiceKey
      );

      setVoiceSetupMessage(
        "Voice Capture key copied."
      );
    } catch (error) {
      if (voiceSetupKey) {
        voiceSetupKey.focus();
        voiceSetupKey.select();
      }

      setVoiceSetupMessage(
        "The key is selected. Choose Copy on your device.",
        true
      );
    }
  }

  function setServerReviewStatus(message, isError) {
    if (!serverReviewStatus) {
      return;
    }

    serverReviewStatus.textContent = message || "";
    serverReviewStatus.classList.toggle(
      "error-message",
      Boolean(isError)
    );
  }

  async function loadServerReviewInbox(forceRefresh) {
    const savedSetup = loadSavedVoiceSetup();

    if (!savedSetup || !savedSetup.voiceKey) {
      setServerReviewStatus(
        "Set up Voice Capture to load spoken captures here."
      );
      return;
    }

    if (!navigator.onLine) {
      setServerReviewStatus(
        "Offline. Showing the last Voice Review items saved on this device."
      );
      return;
    }

    if (serverReviewLoading && !forceRefresh) {
      return;
    }

    serverReviewLoading = true;

    if (refreshReviewButton) {
      refreshReviewButton.disabled = true;
      refreshReviewButton.textContent = "Refreshing...";
    }

    setServerReviewStatus(
      "Loading Voice Capture review items..."
    );

    try {
      const response = await fetch(
        "/api/voice-review-inbox",
        {
          method: "GET",
          headers: {
            Accept: "application/json",
            "x-gsd-capture-key":
              savedSetup.voiceKey,
          },
          cache: "no-store",
        }
      );

      const data = await response.json().catch(function () {
        return null;
      });

      if (!response.ok || !data || !Array.isArray(data.items)) {
        throw new Error(
          data && data.error
            ? data.error
            : "Voice Review items could not be loaded."
        );
      }

      const preservedTasks = tasks.filter(function (task) {
        return !task.serverReviewPending;
      });
      const serverTasks = data.items.map(
        mapServerReviewItemToTask
      );

      tasks = serverTasks.concat(preservedTasks);
      saveTasks();
      renderEverything();

      setServerReviewStatus(
        serverTasks.length === 0
          ? "No new Voice Capture items are waiting for review."
          : serverTasks.length +
              (serverTasks.length === 1
                ? " Voice Capture item is waiting for review."
                : " Voice Capture items are waiting for review.")
      );
    } catch (error) {
      console.error(
        "Voice Review Inbox could not be loaded:",
        error
      );

      setServerReviewStatus(
        error && error.message
          ? error.message
          : "Voice Review items could not be loaded.",
        true
      );
    } finally {
      serverReviewLoading = false;

      if (refreshReviewButton) {
        refreshReviewButton.disabled = false;
        refreshReviewButton.textContent = "Refresh";
      }
    }
  }

  function mapServerReviewItemToTask(item) {
    const reviewId = String(item.id || "");

    return {
      id: "voice-review-" + reviewId,
      serverReviewId: reviewId,
      serverReviewPending: true,
      serverReviewStatus: item.status || "pending",
      source: "voice-server",
      title:
        typeof item.title === "string" && item.title.trim()
          ? item.title.trim()
          : String(item.originalText || "Voice capture"),
      originalText: String(item.originalText || ""),
      category:
        item.suggestedCategory || item.category || "General",
      suggestedCategory:
        item.suggestedCategory || item.category || "General",
      aiConfidence:
        Number.isInteger(item.confidence)
          ? item.confidence
          : null,
      aiExplanation:
        typeof item.explanation === "string"
          ? item.explanation
          : "",
      priority: item.priority || "Normal",
      dueDate: item.dueDate || "",
      context: item.context || "",
      project: item.project || "",
      estimatedMinutes:
        typeof item.estimatedMinutes === "number"
          ? item.estimatedMinutes
          : null,
      needsReview: true,
      sortingMethod: item.sortingMethod || "ai",
      plannedForToday: false,
      completed: false,
      createdAt:
        item.createdAt || new Date().toISOString(),
      microsoftTaskId: item.microsoftTaskId || "",
      microsoftListId:
        item.microsoftReviewListId || "",
      microsoftListName:
        item.microsoftReviewListName ||
        MICROSOFT_REVIEW_LIST_NAME,
      microsoftSyncStatus: "synced",
      microsoftSyncError: "",
      microsoftSyncAttemptedAt: "",
      microsoftSyncedAt:
        item.createdAt || new Date().toISOString(),
      microsoftReviewTaskId:
        item.microsoftTaskId || "",
      microsoftReviewListId:
        item.microsoftReviewListId || "",
      microsoftReviewListName:
        item.microsoftReviewListName ||
        MICROSOFT_REVIEW_LIST_NAME,
      microsoftMoveTargetTaskId: "",
      microsoftMoveTargetListId: "",
      microsoftMoveTargetListName: "",
      microsoftPendingAction: "",
      microsoftActionError: "",
      microsoftActionSyncedAt: "",
      localDeleted: false,
    };
  }

  async function updateServerReviewStatus(
    task,
    status,
    finalListName
  ) {
    if (!task || !task.serverReviewId) {
      return true;
    }

    const savedSetup = loadSavedVoiceSetup();

    if (!savedSetup || !savedSetup.voiceKey) {
      throw new Error(
        "Voice Capture needs to be set up again before this review decision can be saved."
      );
    }

    const response = await fetch(
      "/api/voice-review-inbox",
      {
        method: "PATCH",
        headers: {
          Accept: "application/json",
          "Content-Type": "application/json",
          "x-gsd-capture-key":
            savedSetup.voiceKey,
        },
        body: JSON.stringify({
          reviewId: task.serverReviewId,
          status,
          finalCategory: task.category || "",
          finalListName: finalListName || "",
        }),
      }
    );

    const data = await response.json().catch(function () {
      return null;
    });

    if (!response.ok) {
      throw new Error(
        data && data.error
          ? data.error
          : "The review decision could not be saved."
      );
    }

    task.serverReviewPending = false;
    task.serverReviewStatus = status;
    return true;
  }

  async function keepTaskInGsdReview(task) {
    try {
      await updateServerReviewStatus(
        task,
        "kept",
        MICROSOFT_REVIEW_LIST_NAME
      );

      task.needsReview = false;
      task.microsoftSyncStatus = "synced";
      saveTasks();
      renderEverything();
      showCaptureMessage(
        "Kept in GSD Review. The item is no longer waiting for a decision in the app."
      );
    } catch (error) {
      showCaptureMessage(
        error && error.message
          ? error.message
          : "The review decision could not be saved.",
        true
      );
    }
  }

  async function loadMicrosoftTodoLists() {
    if (!navigator.onLine) {
      setMicrosoftListMessage(
        "Connect to the internet before loading Microsoft To Do lists.",
        true
      );
      return;
    }

    if (!microsoftAuthReady || !microsoftAuth || !microsoftAccount) {
      setMicrosoftListMessage(
        "Connect your Microsoft account first.",
        true
      );
      return;
    }

    setMicrosoftListButtonBusy(true);
    setMicrosoftListMessage(
      "Loading your Microsoft To Do lists..."
    );

    try {
      const accessToken = await acquireMicrosoftAccessToken();

      if (!accessToken) {
        return;
      }

      const controller = new AbortController();
      const timeoutId = window.setTimeout(function () {
        controller.abort();
      }, 15000);

      let response;

      try {
        response = await fetch(
          "https://graph.microsoft.com/v1.0/me/todo/lists",
          {
            headers: {
              Authorization: "Bearer " + accessToken,
              Accept: "application/json",
            },
            signal: controller.signal,
          }
        );
      } finally {
        window.clearTimeout(timeoutId);
      }

      if (!response.ok) {
        throw new Error(
          "Microsoft Graph returned " + response.status + "."
        );
      }

      const data = await response.json();
      microsoftTodoLists = Array.isArray(data.value)
        ? data.value.slice()
        : [];

      microsoftTodoLists.sort(function (first, second) {
        return String(first.displayName || "").localeCompare(
          String(second.displayName || "")
        );
      });

      populateMicrosoftListOptions();
    } catch (error) {
      console.error("Loading Microsoft To Do lists failed:", error);

      const message =
        error && error.name === "AbortError"
          ? "Microsoft To Do took too long to respond. Please try again."
          : "Your Microsoft To Do lists could not be loaded. Please try again.";

      setMicrosoftListMessage(message, true);
    } finally {
      setMicrosoftListButtonBusy(false);
    }
  }

  async function acquireMicrosoftAccessToken() {
    const account =
      microsoftAccount || microsoftAuth.getActiveAccount();

    if (!account) {
      throw new Error("No Microsoft account is connected.");
    }

    try {
      const tokenResult = await microsoftAuth.acquireTokenSilent({
        account: account,
        scopes: MICROSOFT_SCOPES,
      });

      return tokenResult.accessToken;
    } catch (error) {
      const interactionRequired =
        window.msal.InteractionRequiredAuthError &&
        error instanceof window.msal.InteractionRequiredAuthError;

      if (!interactionRequired) {
        throw error;
      }

      setMicrosoftListMessage(
        "Microsoft needs you to approve access again. Opening sign-in..."
      );

      await microsoftAuth.acquireTokenRedirect({
        account: account,
        scopes: MICROSOFT_SCOPES,
      });

      return null;
    }
  }

  function populateMicrosoftListOptions() {
    if (!microsoftListSelect) {
      return;
    }

    microsoftListSelect.innerHTML = "";

    if (microsoftTodoLists.length === 0) {
      const emptyOption = document.createElement("option");
      emptyOption.value = "";
      emptyOption.textContent = "No Microsoft To Do lists found";
      microsoftListSelect.appendChild(emptyOption);
      microsoftListSelect.disabled = true;

      if (createMicrosoftTestTaskButton) {
        createMicrosoftTestTaskButton.disabled = true;
      }

      setMicrosoftListMessage(
        "Microsoft returned no To Do lists for this account.",
        true
      );
      setMicrosoftTestTaskMessage(
        "A destination list is required before creating a test task.",
        true
      );
      return;
    }

    microsoftTodoLists.forEach(function (list) {
      const option = document.createElement("option");
      option.value = list.id;
      option.textContent = list.displayName || "Unnamed list";
      microsoftListSelect.appendChild(option);
    });

    const savedList = microsoftTodoLists.find(function (list) {
      return list.id === settings.microsoftListId;
    });
    const defaultList = microsoftTodoLists.find(function (list) {
      return list.wellknownListName === "defaultList";
    });
    const selectedList =
      savedList || defaultList || microsoftTodoLists[0];

    microsoftListSelect.value = selectedList.id;
    microsoftListSelect.disabled = false;
    saveMicrosoftListSelection();

    if (createMicrosoftTestTaskButton) {
      createMicrosoftTestTaskButton.disabled = false;
    }

    setMicrosoftListMessage(
      "Loaded " +
        microsoftTodoLists.length +
        " To Do " +
        (microsoftTodoLists.length === 1 ? "list" : "lists") +
        ". Destination: " +
        (selectedList.displayName || "Unnamed list") +
        "."
    );
  }

  function saveMicrosoftListSelection() {
    if (!microsoftListSelect || !microsoftListSelect.value) {
      return;
    }

    const selectedList = microsoftTodoLists.find(function (list) {
      return list.id === microsoftListSelect.value;
    });

    if (!selectedList) {
      return;
    }

    settings.microsoftListId = selectedList.id;
    settings.microsoftListName =
      selectedList.displayName || "Unnamed list";

    localStorage.setItem(
      SETTINGS_KEY,
      JSON.stringify(settings)
    );

    setMicrosoftListMessage(
      "Destination list saved: " +
        settings.microsoftListName +
        "."
    );

    if (createMicrosoftTestTaskButton) {
      createMicrosoftTestTaskButton.disabled = false;
    }

    setMicrosoftTestTaskMessage(
      "Every capture will sync to Microsoft To Do. Unclear items go to GSD Review until you approve them. Approved items use " +
        settings.microsoftListName +
        " or an AI-routed list."
    );
  }

  function saveMicrosoftRoutingMode() {
    if (!microsoftRoutingModeSelect) {
      return;
    }

    settings.aiListRouting = microsoftRoutingModeSelect.value || "ask";

    localStorage.setItem(
      SETTINGS_KEY,
      JSON.stringify(settings)
    );

    updateMicrosoftRoutingMessage();
  }

  function updateMicrosoftRoutingMessage() {
    if (!microsoftRoutingMessage) {
      return;
    }

    const mode = settings.aiListRouting || "ask";

    if (mode === "automatic") {
      microsoftRoutingMessage.textContent =
        "AI categories will automatically create and use matching To Do lists.";
      return;
    }

    if (mode === "selected-only") {
      microsoftRoutingMessage.textContent =
        "Every capture will use the selected destination list.";
      return;
    }

    microsoftRoutingMessage.textContent =
      "Every unclear capture goes to GSD Review. When an approved AI category has no matching list, GSD Capture will ask before creating it.";
  }

  function normalizeMicrosoftListName(value) {
    return String(value || "")
      .trim()
      .toLocaleLowerCase();
  }

  function getSuggestedMicrosoftListName(task) {
    const category = String(
      task && task.approvedListName
        ? task.approvedListName
        : task && task.category
          ? task.category
          : ""
    ).trim();

    if (!category || category.toLocaleLowerCase() === "general") {
      return "";
    }

    return category.slice(0, 80);
  }

  function findMicrosoftListByName(displayName) {
    const normalizedName = normalizeMicrosoftListName(displayName);

    return microsoftTodoLists.find(function (list) {
      return (
        normalizeMicrosoftListName(list.displayName) === normalizedName
      );
    });
  }

  async function ensureMicrosoftTodoListsForRouting(accessToken) {
    if (microsoftTodoLists.length > 0) {
      return microsoftTodoLists;
    }

    const response = await fetch(
      "https://graph.microsoft.com/v1.0/me/todo/lists",
      {
        headers: {
          Authorization: "Bearer " + accessToken,
          Accept: "application/json",
        },
      }
    );

    const data = await response.json().catch(function () {
      return null;
    });

    if (!response.ok) {
      const graphMessage =
        data && data.error && data.error.message
          ? data.error.message
          : "Microsoft Graph returned " + response.status + ".";

      throw new Error(graphMessage);
    }

    microsoftTodoLists =
      data && Array.isArray(data.value) ? data.value.slice() : [];

    microsoftTodoLists.sort(function (first, second) {
      return String(first.displayName || "").localeCompare(
        String(second.displayName || "")
      );
    });

    return microsoftTodoLists;
  }

  async function createMicrosoftTodoList(displayName, accessToken) {
    const response = await fetch(
      "https://graph.microsoft.com/v1.0/me/todo/lists",
      {
        method: "POST",
        headers: {
          Authorization: "Bearer " + accessToken,
          Accept: "application/json",
          "Content-Type": "application/json",
        },
        body: JSON.stringify({
          displayName: displayName,
        }),
      }
    );

    const data = await response.json().catch(function () {
      return null;
    });

    if (!response.ok) {
      const graphMessage =
        data && data.error && data.error.message
          ? data.error.message
          : "Microsoft Graph returned " + response.status + ".";

      throw new Error(graphMessage);
    }

    if (!data || !data.id) {
      throw new Error(
        "Microsoft To Do did not return a list ID."
      );
    }

    microsoftTodoLists.push(data);
    microsoftTodoLists.sort(function (first, second) {
      return String(first.displayName || "").localeCompare(
        String(second.displayName || "")
      );
    });

    return data;
  }

  async function getOrCreateMicrosoftListByName(displayName, accessToken) {
    await ensureMicrosoftTodoListsForRouting(accessToken);

    let matchingList = findMicrosoftListByName(displayName);

    if (!matchingList) {
      matchingList = await createMicrosoftTodoList(
        displayName,
        accessToken
      );
    }

    return matchingList;
  }

  async function routeTaskToMicrosoftList(task, accessToken, options) {
    const routeOptions = options || {};
    const routeAsApproved = Boolean(routeOptions.routeAsApproved);

    if (task.microsoftListId && !routeAsApproved) {
      return { listId: task.microsoftListId, created: false };
    }

    if (task.needsReview && !routeAsApproved) {
      const reviewList = await getOrCreateMicrosoftListByName(
        MICROSOFT_REVIEW_LIST_NAME,
        accessToken
      );

      task.microsoftListId = reviewList.id;
      task.microsoftListName =
        reviewList.displayName || MICROSOFT_REVIEW_LIST_NAME;
      task.microsoftReviewListId = reviewList.id;
      task.microsoftReviewListName =
        reviewList.displayName || MICROSOFT_REVIEW_LIST_NAME;

      return {
        listId: reviewList.id,
        created: false,
        reviewList: true,
      };
    }

    const selectedListId = settings.microsoftListId || "";
    const selectedListName =
      settings.microsoftListName || "Microsoft To Do";
    const mode = settings.aiListRouting || "ask";
    const suggestedListName = getSuggestedMicrosoftListName(task);

    if (task.approvedListName && suggestedListName) {
      await ensureMicrosoftTodoListsForRouting(accessToken);

      let approvedList = findMicrosoftListByName(
        suggestedListName
      );

      if (!approvedList) {
        approvedList = await createMicrosoftTodoList(
          suggestedListName,
          accessToken
        );
      }

      task.microsoftListId = approvedList.id;
      task.microsoftListName =
        approvedList.displayName || suggestedListName;

      return {
        listId: approvedList.id,
        created: false,
      };
    }

    if (mode === "selected-only" || !suggestedListName) {
      task.microsoftListId = selectedListId;
      task.microsoftListName = selectedListName;
      return { listId: selectedListId, created: false };
    }

    try {
      await ensureMicrosoftTodoListsForRouting(accessToken);

      let matchingList = findMicrosoftListByName(suggestedListName);

      if (!matchingList) {
        let shouldCreate = mode === "automatic";

        if (mode === "ask") {
          shouldCreate = window.confirm(
            'AI suggests the Microsoft To Do list "' +
              suggestedListName +
              '". Create it and add this task there?'
          );
        }

        if (shouldCreate) {
          matchingList = await createMicrosoftTodoList(
            suggestedListName,
            accessToken
          );
          task.microsoftListCreatedByAi = true;
        }
      }

      if (matchingList) {
        task.microsoftListId = matchingList.id;
        task.microsoftListName =
          matchingList.displayName || suggestedListName;
        return {
          listId: matchingList.id,
          created: Boolean(task.microsoftListCreatedByAi),
        };
      }
    } catch (error) {
      console.warn(
        "AI list routing could not finish; using the selected list instead:",
        error
      );
    }

    task.microsoftListId = selectedListId;
    task.microsoftListName = selectedListName;
    return { listId: selectedListId, created: false };
  }

  async function createMicrosoftTestTask() {
    if (!navigator.onLine) {
      setMicrosoftTestTaskMessage(
        "Connect to the internet before creating a Microsoft To Do task.",
        true
      );
      return;
    }

    if (!microsoftAuthReady || !microsoftAuth || !microsoftAccount) {
      setMicrosoftTestTaskMessage(
        "Connect your Microsoft account first.",
        true
      );
      return;
    }

    if (!settings.microsoftListId) {
      setMicrosoftTestTaskMessage(
        "Load your To Do lists and choose a destination first.",
        true
      );
      return;
    }

    setMicrosoftTestTaskButtonBusy(true);
    setMicrosoftTestTaskMessage(
      "Creating a test task in " +
        (settings.microsoftListName || "Microsoft To Do") +
        "..."
    );

    try {
      const accessToken = await acquireMicrosoftAccessToken();

      if (!accessToken) {
        return;
      }

      const controller = new AbortController();
      const timeoutId = window.setTimeout(function () {
        controller.abort();
      }, 15000);

      let response;

      try {
        response = await fetch(
          "https://graph.microsoft.com/v1.0/me/todo/lists/" +
            encodeURIComponent(settings.microsoftListId) +
            "/tasks",
          {
            method: "POST",
            headers: {
              Authorization: "Bearer " + accessToken,
              Accept: "application/json",
              "Content-Type": "application/json",
            },
            body: JSON.stringify({
              title: "GSD Capture connection test",
              body: {
                contentType: "text",
                content:
                  "Created by GSD Capture to verify Microsoft To Do integration on " +
                  new Date().toLocaleString() +
                  ".",
              },
            }),
            signal: controller.signal,
          }
        );
      } finally {
        window.clearTimeout(timeoutId);
      }

      const responseData = await response.json().catch(function () {
        return null;
      });

      if (!response.ok) {
        const graphMessage =
          responseData &&
          responseData.error &&
          responseData.error.message
            ? responseData.error.message
            : "Microsoft Graph returned " + response.status + ".";

        throw new Error(graphMessage);
      }

      setMicrosoftTestTaskMessage(
        'Created "' +
          (responseData && responseData.title
            ? responseData.title
            : "GSD Capture connection test") +
          '" in ' +
          (settings.microsoftListName || "Microsoft To Do") +
          "."
      );
    } catch (error) {
      console.error("Creating Microsoft To Do test task failed:", error);

      const message =
        error && error.name === "AbortError"
          ? "Microsoft To Do took too long to respond. Please try again."
          : "The test task could not be created. Please try again.";

      setMicrosoftTestTaskMessage(message, true);
    } finally {
      setMicrosoftTestTaskButtonBusy(false);
    }
  }

  function canSyncTaskToMicrosoft(task) {
    return Boolean(
      navigator.onLine &&
      microsoftAuthReady &&
      microsoftAuth &&
      microsoftAccount &&
      ((task && task.microsoftListId) || settings.microsoftListId)
    );
  }

  function getMicrosoftListIdForTask(task) {
    return task.microsoftListId || settings.microsoftListId;
  }

  function getMicrosoftListNameForTask(task) {
    return (
      task.microsoftListName ||
      settings.microsoftListName ||
      "Microsoft To Do"
    );
  }

  async function syncTaskToMicrosoft(task) {
    if (!task) {
      return { status: "failed" };
    }

    if (task.microsoftTaskId) {
      task.microsoftSyncStatus = "synced";
      saveTasks();
      renderEverything();
      return { status: "synced", taskId: task.microsoftTaskId };
    }

    if (!canSyncTaskToMicrosoft(task)) {
      task.microsoftSyncStatus = "local-only";
      task.microsoftSyncError =
        "Microsoft To Do is not connected or no destination list is selected.";
      saveTasks();
      renderEverything();
      return { status: "local-only" };
    }

    const hadPreviousAttempt = Boolean(
      task.microsoftSyncAttemptedAt
    );

    task.microsoftSyncStatus = "syncing";
    task.microsoftSyncError = "";
    task.microsoftListId = task.microsoftListId || "";
    task.microsoftListName = task.microsoftListName || "";
    task.microsoftSyncAttemptedAt = new Date().toISOString();
    saveTasks();
    renderEverything();

    try {
      const accessToken = await acquireMicrosoftAccessToken();

      if (!accessToken) {
        task.microsoftSyncStatus = "pending";
        saveTasks();
        renderEverything();
        return { status: "interaction-required" };
      }

      await routeTaskToMicrosoftList(task, accessToken);
      saveTasks();
      renderEverything();

      if (hadPreviousAttempt) {
        const existingTask =
          await findExistingMicrosoftTask(task, accessToken);

        if (existingTask && existingTask.id) {
          markTaskMicrosoftSynced(task, existingTask);
          return {
            status: "synced",
            taskId: existingTask.id,
            recovered: true,
          };
        }
      }

      const responseData = await createMicrosoftTask(
        task,
        accessToken
      );

      markTaskMicrosoftSynced(task, responseData);

      return {
        status: "synced",
        taskId: responseData.id,
      };
    } catch (error) {
      console.error("Microsoft To Do task sync failed:", error);

      task.microsoftSyncStatus = "failed";
      task.microsoftSyncError =
        error && error.message
          ? error.message
          : "Microsoft To Do sync failed.";
      saveTasks();
      renderEverything();

      return { status: "failed", error: error };
    }
  }

  async function createMicrosoftTask(task, accessToken, listIdOverride) {
    const controller = new AbortController();
    const timeoutId = window.setTimeout(function () {
      controller.abort();
    }, 15000);

    let response;

    try {
      response = await fetch(
        "https://graph.microsoft.com/v1.0/me/todo/lists/" +
          encodeURIComponent(
            listIdOverride || getMicrosoftListIdForTask(task)
          ) +
          "/tasks",
        {
          method: "POST",
          headers: {
            Authorization: "Bearer " + accessToken,
            Accept: "application/json",
            "Content-Type": "application/json",
          },
          body: JSON.stringify(
            buildMicrosoftTaskPayload(task)
          ),
          signal: controller.signal,
        }
      );
    } finally {
      window.clearTimeout(timeoutId);
    }

    const responseData = await response.json().catch(function () {
      return null;
    });

    if (!response.ok) {
      const graphMessage =
        responseData &&
        responseData.error &&
        responseData.error.message
          ? responseData.error.message
          : "Microsoft Graph returned " + response.status + ".";

      throw new Error(graphMessage);
    }

    if (!responseData || !responseData.id) {
      throw new Error(
        "Microsoft To Do did not return a task ID."
      );
    }

    return responseData;
  }

  function buildMicrosoftTaskPayload(task) {
    const notes = [];

    if (task.originalText && task.originalText !== task.title) {
      notes.push("Original capture: " + task.originalText);
    }

    notes.push("Category: " + (task.category || "General"));
    notes.push("Priority: " + (task.priority || "Normal"));

    if (task.needsReview) {
      notes.push("Review status: Needs review in GSD Capture");
    }

    if (task.context) {
      notes.push("Context: " + task.context);
    }

    if (task.project) {
      notes.push("Project: " + task.project);
    }

    if (task.estimatedMinutes) {
      notes.push(
        "Estimated duration: " +
          task.estimatedMinutes +
          " minutes"
      );
    }

    notes.push(
      "Sorted by: " +
        (task.sortingMethod === "ai"
          ? "GSD Capture AI"
          : "GSD Capture local sorting")
    );
    notes.push(getMicrosoftTaskMarker(task));

    const payload = {
      title: task.title,
      importance: getMicrosoftImportance(task.priority),
      body: {
        contentType: "text",
        content: notes.join("\n"),
      },
    };

    if (task.dueDate) {
      payload.dueDateTime = {
        dateTime: task.dueDate + "T12:00:00",
        timeZone: "UTC",
      };
    }

    return payload;
  }

  function getMicrosoftImportance(priority) {
    if (priority === "High") {
      return "high";
    }

    if (priority === "Low") {
      return "low";
    }

    return "normal";
  }

  function getMicrosoftTaskMarker(task) {
    return "GSD Capture ID: " + task.id;
  }

  async function findExistingMicrosoftTask(task, accessToken, listIdOverride) {
    let nextUrl =
      "https://graph.microsoft.com/v1.0/me/todo/lists/" +
      encodeURIComponent(
        listIdOverride || getMicrosoftListIdForTask(task)
      ) +
      "/tasks";
    const marker = getMicrosoftTaskMarker(task);
    let pageCount = 0;

    while (nextUrl && pageCount < 20) {
      const controller = new AbortController();
      const timeoutId = window.setTimeout(function () {
        controller.abort();
      }, 15000);
      let response;

      try {
        response = await fetch(nextUrl, {
          headers: {
            Authorization: "Bearer " + accessToken,
            Accept: "application/json",
          },
          signal: controller.signal,
        });
      } finally {
        window.clearTimeout(timeoutId);
      }

      const data = await response.json().catch(function () {
        return null;
      });

      if (!response.ok) {
        const graphMessage =
          data && data.error && data.error.message
            ? data.error.message
            : "Microsoft Graph returned " + response.status + ".";
        throw new Error(graphMessage);
      }

      const remoteTasks =
        data && Array.isArray(data.value) ? data.value : [];

      const match = remoteTasks.find(function (remoteTask) {
        return Boolean(
          remoteTask &&
          remoteTask.body &&
          typeof remoteTask.body.content === "string" &&
          remoteTask.body.content.indexOf(marker) !== -1
        );
      });

      if (match) {
        return match;
      }

      nextUrl =
        data && typeof data["@odata.nextLink"] === "string"
          ? data["@odata.nextLink"]
          : "";
      pageCount += 1;
    }

    return null;
  }

  function markTaskMicrosoftSynced(task, microsoftTask) {
    task.microsoftTaskId = microsoftTask.id;
    task.microsoftListId = getMicrosoftListIdForTask(task);
    task.microsoftListName = getMicrosoftListNameForTask(task);

    if (task.needsReview) {
      task.microsoftReviewTaskId = microsoftTask.id;
      task.microsoftReviewListId = task.microsoftListId;
      task.microsoftReviewListName =
        task.microsoftListName || MICROSOFT_REVIEW_LIST_NAME;
    }

    task.microsoftSyncStatus = "synced";
    task.microsoftSyncError = "";
    task.microsoftSyncedAt = new Date().toISOString();
    saveTasks();
    renderEverything();
  }

  async function deleteMicrosoftTaskFromList(listId, taskId, accessToken) {
    const response = await fetch(
      "https://graph.microsoft.com/v1.0/me/todo/lists/" +
        encodeURIComponent(listId) +
        "/tasks/" +
        encodeURIComponent(taskId),
      {
        method: "DELETE",
        headers: {
          Authorization: "Bearer " + accessToken,
        },
      }
    );

    if (!response.ok && response.status !== 404) {
      const data = await response.json().catch(function () {
        return null;
      });
      const graphMessage =
        data && data.error && data.error.message
          ? data.error.message
          : "Microsoft Graph returned " + response.status + ".";
      throw new Error(graphMessage);
    }
  }


  function hasMicrosoftTaskReference(task) {
    return Boolean(
      task &&
        task.microsoftTaskId &&
        task.microsoftListId
    );
  }

  async function updateMicrosoftTaskCompletion(
    listId,
    taskId,
    accessToken
  ) {
    const controller = new AbortController();
    const timeoutId = window.setTimeout(function () {
      controller.abort();
    }, 15000);

    let response;

    try {
      response = await fetch(
        "https://graph.microsoft.com/v1.0/me/todo/lists/" +
          encodeURIComponent(listId) +
          "/tasks/" +
          encodeURIComponent(taskId),
        {
          method: "PATCH",
          headers: {
            Authorization: "Bearer " + accessToken,
            Accept: "application/json",
            "Content-Type": "application/json",
          },
          body: JSON.stringify({
            status: "completed",
          }),
          signal: controller.signal,
        }
      );
    } finally {
      window.clearTimeout(timeoutId);
    }

    if (!response.ok && response.status !== 404) {
      const data = await response.json().catch(function () {
        return null;
      });
      const graphMessage =
        data && data.error && data.error.message
          ? data.error.message
          : "Microsoft Graph returned " + response.status + ".";
      throw new Error(graphMessage);
    }
  }

  async function processMicrosoftActionForTask(
    task,
    accessTokenOverride
  ) {
    if (!task || !task.microsoftPendingAction) {
      return { status: "none" };
    }

    const action = task.microsoftPendingAction;

    if (!hasMicrosoftTaskReference(task)) {
      task.microsoftPendingAction = "";
      task.microsoftActionError = "";

      if (action === "delete") {
        tasks = tasks.filter(function (item) {
          return item.id !== task.id;
        });
      }

      saveTasks();
      renderEverything();
      return { status: "synced" };
    }

    if (
      !navigator.onLine ||
      !microsoftAuthReady ||
      !microsoftAuth ||
      !microsoftAccount
    ) {
      return { status: "pending" };
    }

    try {
      const accessToken =
        accessTokenOverride ||
        (await acquireMicrosoftAccessToken());

      if (!accessToken) {
        return { status: "interaction-required" };
      }

      if (action === "complete") {
        await updateMicrosoftTaskCompletion(
          task.microsoftListId,
          task.microsoftTaskId,
          accessToken
        );

        task.microsoftPendingAction = "";
        task.microsoftActionError = "";
        task.microsoftActionSyncedAt = new Date().toISOString();
        saveTasks();
        renderEverything();

        return { status: "synced" };
      }

      if (action === "delete") {
        await deleteMicrosoftTaskFromList(
          task.microsoftListId,
          task.microsoftTaskId,
          accessToken
        );

        tasks = tasks.filter(function (item) {
          return item.id !== task.id;
        });
        saveTasks();
        renderEverything();

        return { status: "synced" };
      }

      task.microsoftPendingAction = "";
      task.microsoftActionError = "";
      saveTasks();
      renderEverything();
      return { status: "none" };
    } catch (error) {
      console.error(
        "Microsoft To Do action sync failed:",
        error
      );

      task.microsoftActionError =
        error && error.message
          ? error.message
          : "Microsoft To Do action sync failed.";
      saveTasks();
      renderEverything();

      return { status: "pending", error: error };
    }
  }

  async function processPendingMicrosoftActions() {
    if (
      microsoftActionQueueRunning ||
      !navigator.onLine ||
      !microsoftAuthReady ||
      !microsoftAuth ||
      !microsoftAccount
    ) {
      return;
    }

    const pendingTasks = tasks.filter(function (task) {
      return Boolean(task.microsoftPendingAction);
    });

    if (pendingTasks.length === 0) {
      return;
    }

    microsoftActionQueueRunning = true;

    try {
      const accessToken = await acquireMicrosoftAccessToken();

      if (!accessToken) {
        return;
      }

      for (const pendingTask of pendingTasks) {
        await processMicrosoftActionForTask(
          pendingTask,
          accessToken
        );
      }
    } catch (error) {
      console.error(
        "Processing pending Microsoft To Do actions failed:",
        error
      );
    } finally {
      microsoftActionQueueRunning = false;
    }
  }

  async function completeTaskAndSync(task) {
    if (!task || task.localDeleted || task.completed) {
      return;
    }

    task.completed = true;
    task.needsReview = false;
    task.plannedForToday = false;
    task.microsoftActionError = "";

    if (hasMicrosoftTaskReference(task)) {
      task.microsoftPendingAction = "complete";
    } else {
      task.microsoftPendingAction = "";
    }

    saveTasks();
    renderEverything();

    if (!task.microsoftPendingAction) {
      if (task.serverReviewId) {
        try {
          await updateServerReviewStatus(
            task,
            "completed",
            task.microsoftListName || ""
          );
        } catch (error) {
          console.error(
            "The server Review Inbox could not save completion:",
            error
          );
        }
      }

      showCaptureMessage(
        "Completed in GSD Capture. This item did not have a Microsoft To Do copy."
      );
      return;
    }

    const result = await processMicrosoftActionForTask(task);

    if (result.status === "synced") {
      if (task.serverReviewId) {
        try {
          await updateServerReviewStatus(
            task,
            "completed",
            task.microsoftListName || ""
          );
        } catch (error) {
          console.error(
            "The server Review Inbox could not save completion:",
            error
          );
        }
      }

      showCaptureMessage(
        "Completed in GSD Capture and Microsoft To Do."
      );
      return;
    }

    showCaptureMessage(
      "Completed in GSD Capture. Microsoft To Do will update automatically when the connection is available."
    );
  }

  async function deleteTaskAndSync(task) {
    if (!task) {
      return;
    }

    if (!hasMicrosoftTaskReference(task)) {
      if (task.serverReviewId) {
        try {
          await updateServerReviewStatus(
            task,
            "rejected",
            ""
          );
        } catch (error) {
          showCaptureMessage(
            "The item could not be removed from the server Review Inbox.",
            true
          );
          return;
        }
      }

      tasks = tasks.filter(function (item) {
        return item.id !== task.id;
      });
      saveTasks();
      renderEverything();
      showCaptureMessage("Deleted from GSD Capture.");
      return;
    }

    task.localDeleted = true;
    task.microsoftPendingAction = "delete";
    task.microsoftActionError = "";
    saveTasks();
    renderEverything();

    const result = await processMicrosoftActionForTask(task);

    if (result.status === "synced") {
      if (task.serverReviewId) {
        try {
          await updateServerReviewStatus(
            task,
            "rejected",
            ""
          );
        } catch (error) {
          console.error(
            "The server Review Inbox could not save the rejection:",
            error
          );
        }
      }

      showCaptureMessage(
        "Deleted from GSD Capture and Microsoft To Do."
      );
      return;
    }

    showCaptureMessage(
      "Hidden from GSD Capture. Microsoft To Do will delete it automatically when the connection is available."
    );
  }

  async function moveReviewedTaskToDestination(task) {
    if (!canSyncTaskToMicrosoft(task)) {
      task.microsoftSyncStatus = "failed";
      task.microsoftSyncError =
        "Connect Microsoft To Do and select a destination list before approving this item.";
      saveTasks();
      renderEverything();
      return { status: "failed" };
    }

    task.microsoftSyncStatus = "moving";
    task.microsoftSyncError = "";
    saveTasks();
    renderEverything();

    try {
      const accessToken = await acquireMicrosoftAccessToken();

      if (!accessToken) {
        task.microsoftSyncStatus = "pending";
        saveTasks();
        renderEverything();
        return { status: "interaction-required" };
      }

      const reviewTaskId =
        task.microsoftReviewTaskId || task.microsoftTaskId;
      const reviewListId =
        task.microsoftReviewListId || task.microsoftListId;

      if (!reviewTaskId || !reviewListId) {
        task.needsReview = false;
        task.microsoftTaskId = "";
        task.microsoftListId = "";
        task.microsoftListName = "";
        task.microsoftSyncStatus = "pending";
        saveTasks();
        renderEverything();
        return await syncTaskToMicrosoft(task);
      }

      if (!task.microsoftMoveTargetTaskId) {
        const originalNeedsReview = task.needsReview;
        const originalTaskId = task.microsoftTaskId;
        const originalListId = task.microsoftListId;
        const originalListName = task.microsoftListName;

        task.needsReview = false;
        task.microsoftTaskId = "";
        task.microsoftListId = "";
        task.microsoftListName = "";

        await routeTaskToMicrosoftList(task, accessToken, {
          routeAsApproved: true,
        });

        const destinationListId = task.microsoftListId;
        const destinationListName = task.microsoftListName;
        const existingTask = await findExistingMicrosoftTask(
          task,
          accessToken,
          destinationListId
        );
        const destinationTask =
          existingTask && existingTask.id
            ? existingTask
            : await createMicrosoftTask(
                task,
                accessToken,
                destinationListId
              );

        task.microsoftMoveTargetTaskId = destinationTask.id;
        task.microsoftMoveTargetListId = destinationListId;
        task.microsoftMoveTargetListName = destinationListName;

        task.needsReview = originalNeedsReview;
        task.microsoftTaskId = originalTaskId;
        task.microsoftListId = originalListId;
        task.microsoftListName = originalListName;
        saveTasks();
      }

      await deleteMicrosoftTaskFromList(
        reviewListId,
        reviewTaskId,
        accessToken
      );

      task.needsReview = false;
      task.microsoftTaskId = task.microsoftMoveTargetTaskId;
      task.microsoftListId = task.microsoftMoveTargetListId;
      task.microsoftListName = task.microsoftMoveTargetListName;
      task.microsoftSyncStatus = "synced";
      task.microsoftSyncError = "";
      task.microsoftSyncedAt = new Date().toISOString();
      task.microsoftReviewTaskId = "";
      task.microsoftReviewListId = "";
      task.microsoftReviewListName = "";
      task.microsoftMoveTargetTaskId = "";
      task.microsoftMoveTargetListId = "";
      task.microsoftMoveTargetListName = "";
      saveTasks();
      renderEverything();

      return { status: "synced", taskId: task.microsoftTaskId };
    } catch (error) {
      console.error("Moving reviewed Microsoft To Do task failed:", error);
      task.microsoftSyncStatus = "failed";
      task.microsoftSyncError =
        error && error.message
          ? error.message
          : "The reviewed task could not be moved.";
      saveTasks();
      renderEverything();
      return { status: "failed", error: error };
    }
  }

  async function chooseDifferentFolderAndApprove(task) {
    const folderName = window.prompt(
      "Enter the Microsoft To Do folder name for this task:",
      task.category || task.suggestedCategory || "General"
    );

    if (!folderName || !folderName.trim()) {
      return;
    }

    task.category = folderName.trim().slice(0, 80);
    task.approvedListName = task.category;
    saveTasks();
    renderEverything();

    await approveTaskAndSync(task);
  }

  async function approveTaskAndSync(task) {
    const hasReviewTask = Boolean(
      task.microsoftReviewTaskId ||
        (task.microsoftTaskId &&
          normalizeMicrosoftListName(task.microsoftListName) ===
            normalizeMicrosoftListName(MICROSOFT_REVIEW_LIST_NAME))
    );

    if (hasReviewTask) {
      const result = await moveReviewedTaskToDestination(task);

      if (result.status === "synced") {
        try {
          await updateServerReviewStatus(
            task,
            "approved",
            task.microsoftListName || task.category
          );
          saveTasks();
          renderEverything();
          showCaptureMessage(
            "Approved and moved to " +
              (task.microsoftListName || task.category || "Microsoft To Do") +
              "."
          );
        } catch (error) {
          console.error(
            "The server review decision could not be saved:",
            error
          );
          showCaptureMessage(
            "The task moved in Microsoft To Do, but the app could not save the review decision. Refresh and try again.",
            true
          );
        }
      }

      return;
    }

    task.needsReview = false;
    task.microsoftSyncStatus = "pending";
    saveTasks();
    renderEverything();

    if (!canSyncTaskToMicrosoft(task)) {
      task.microsoftSyncStatus = "local-only";
      saveTasks();
      renderEverything();
      return;
    }

    const result = await syncTaskToMicrosoft(task);

    if (result.status === "synced") {
      await updateServerReviewStatus(
        task,
        "approved",
        task.microsoftListName || task.category
      );
      saveTasks();
      renderEverything();
    }
  }

  function setMicrosoftTestTaskButtonBusy(isBusy) {
    if (!createMicrosoftTestTaskButton) {
      return;
    }

    createMicrosoftTestTaskButton.disabled = isBusy;
    createMicrosoftTestTaskButton.textContent = isBusy
      ? "Creating Test Task..."
      : "Create Test Task";
  }

  function setMicrosoftTestTaskMessage(message, isError) {
    if (!microsoftTestTaskMessage) {
      return;
    }

    microsoftTestTaskMessage.textContent = message;
    microsoftTestTaskMessage.classList.toggle(
      "error-message",
      Boolean(isError)
    );
  }

  function clearMicrosoftListOptions() {
    microsoftTodoLists = [];

    if (microsoftListSelect) {
      microsoftListSelect.innerHTML =
        '<option value="">Load your Microsoft To Do lists</option>';
      microsoftListSelect.disabled = true;
    }

    if (createMicrosoftTestTaskButton) {
      createMicrosoftTestTaskButton.disabled = true;
    }

    setMicrosoftListMessage(
      "No Microsoft To Do lists loaded yet."
    );
    setMicrosoftTestTaskMessage(
      "Load your lists and choose a destination before creating a test task."
    );
  }

  function setMicrosoftListButtonBusy(isBusy) {
    if (!loadMicrosoftListsButton) {
      return;
    }

    loadMicrosoftListsButton.disabled = isBusy;
    loadMicrosoftListsButton.textContent = isBusy
      ? "Loading Lists..."
      : "Load To Do Lists";
  }

  function setMicrosoftListMessage(message, isError) {
    if (!microsoftListMessage) {
      return;
    }

    microsoftListMessage.textContent = message;
    microsoftListMessage.classList.toggle(
      "error-message",
      Boolean(isError)
    );
  }

  function setMicrosoftMessage(message, isError) {
    if (!microsoftStatus) {
      return;
    }

    microsoftStatus.textContent = message;
    microsoftStatus.classList.toggle(
      "error-message",
      Boolean(isError)
    );
  }

  async function captureTask(event) {
    event.preventDefault();

    if (captureInProgress) {
      return;
    }

    const text = taskInput.value.trim();

    if (!text) {
      showCaptureMessage(
        "Enter a task or idea first.",
        true
      );
      return;
    }

    captureInProgress = true;
    setCaptureButtonBusy(true);

    try {
      let analysis = null;
      let usedAi = false;

      if (settings.aiSorting !== false && navigator.onLine) {
        showCaptureMessage(
          "Capturing and sorting with AI..."
        );

        try {
          analysis = await classifyTaskWithAi(text);
          usedAi = true;
        } catch (aiError) {
          console.error("AI sorting failed:", aiError);
        }
      }

      if (!analysis) {
        analysis = analyzeTask(text);
      }

      const newTask = {
        id:
          Date.now().toString() +
          "-" +
          Math.random().toString(16).slice(2),
        title:
          analysis.summary &&
          analysis.summary.trim()
            ? analysis.summary.trim()
            : text,
        originalText: text,
        category: analysis.category || "General",
        priority: analysis.priority || "Normal",
        dueDate: analysis.dueDate || "",
        context: analysis.context || "",
        project: analysis.project || "",
        estimatedMinutes:
          typeof analysis.estimatedMinutes === "number"
            ? analysis.estimatedMinutes
            : null,
        needsReview:
          typeof analysis.needsReview === "boolean"
            ? analysis.needsReview
            : true,
        sortingMethod: usedAi ? "ai" : "local",
        plannedForToday: false,
        completed: false,
        createdAt: new Date().toISOString(),
        microsoftTaskId: "",
        microsoftListId: "",
        microsoftListName: "",
        microsoftSyncStatus: "pending",
        microsoftSyncError: "",
        microsoftSyncAttemptedAt: "",
        microsoftSyncedAt: "",
        microsoftReviewTaskId: "",
        microsoftReviewListId: "",
        microsoftReviewListName: "",
        microsoftMoveTargetTaskId: "",
        microsoftMoveTargetListId: "",
        microsoftMoveTargetListName: "",
        microsoftPendingAction: "",
        microsoftActionError: "",
        microsoftActionSyncedAt: "",
        localDeleted: false,
      };

      tasks.unshift(newTask);
      saveTasks();

      taskInput.value = "";
      renderEverything();

      const baseMessage = usedAi
        ? newTask.needsReview
          ? "Captured and sorted with AI. This item needs review."
          : "Captured and sorted with AI."
        : navigator.onLine
          ? "Captured successfully using local sorting."
          : "Captured offline using local sorting.";

      if (!canSyncTaskToMicrosoft(newTask)) {
        newTask.microsoftSyncStatus = "local-only";
        saveTasks();
        renderEverything();

        showCaptureMessage(
          baseMessage +
            (newTask.needsReview
              ? " This item needs review and was saved locally. Connect Microsoft To Do to place it in GSD Review."
              : " Saved locally. Connect Microsoft To Do and select a list to send it later.")
        );
        return;
      }

      showCaptureMessage(
        baseMessage + " Sending to Microsoft To Do..."
      );

      const syncResult = await syncTaskToMicrosoft(newTask);

      if (syncResult.status === "synced") {
        showCaptureMessage(
          baseMessage +
            (newTask.needsReview
              ? " Added to GSD Review in Microsoft To Do. Approve it in GSD Capture to move it to the correct list."
              : " Added to " +
                (newTask.microsoftListName ||
                  settings.microsoftListName ||
                  "Microsoft To Do") +
                ".")
        );
      } else if (syncResult.status === "interaction-required") {
        showCaptureMessage(
          baseMessage +
            " Saved locally. Microsoft needs you to finish signing in before it can sync."
        );
      } else {
        showCaptureMessage(
          baseMessage +
            " Saved locally, but Microsoft To Do sync did not finish. Use Send to To Do on the task to retry."
        );
      }
    } catch (error) {
      console.error(error);

      showCaptureMessage(
        "Capture error: " + error.message,
        true
      );
    } finally {
      captureInProgress = false;
      setCaptureButtonBusy(false);
    }
  }

  async function classifyTaskWithAi(text) {
    const controller = new AbortController();
    const timeoutId = window.setTimeout(function () {
      controller.abort();
    }, 15000);

    try {
      const apiResponse = await fetch("/api/classify", {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
        },
        body: JSON.stringify({
          text: text,
          currentDate: getDateString(new Date()),
        }),
        signal: controller.signal,
      });

      const responseData = await apiResponse.json();

      if (!apiResponse.ok) {
        throw new Error(
          responseData && responseData.error
            ? responseData.error
            : "AI sorting was unavailable."
        );
      }

      if (
        !responseData ||
        !responseData.classification
      ) {
        throw new Error(
          "The AI response did not contain a classification."
        );
      }

      return normalizeClassification(
        responseData.classification
      );
    } finally {
      window.clearTimeout(timeoutId);
    }
  }

  function normalizeClassification(classification) {
    const allowedCategories = [
      "Purchasing",
      "Operations",
      "Leadership",
      "Personal",
      "Ideas",
      "General",
    ];

    const allowedPriorities = [
      "High",
      "Normal",
      "Low",
    ];

    const allowedContexts = [
      "Calls",
      "Computer",
      "Errands",
      "Work",
      "Home",
      "Anywhere",
    ];

    return {
      summary:
        typeof classification.summary === "string"
          ? classification.summary
          : "",
      category:
        allowedCategories.indexOf(
          classification.category
        ) >= 0
          ? classification.category
          : "General",
      priority:
        allowedPriorities.indexOf(
          classification.priority
        ) >= 0
          ? classification.priority
          : "Normal",
      dueDate:
        typeof classification.dueDate === "string"
          ? classification.dueDate
          : "",
      context:
        allowedContexts.indexOf(
          classification.context
        ) >= 0
          ? classification.context
          : "",
      project:
        typeof classification.project === "string"
          ? classification.project
          : "",
      estimatedMinutes:
        Number.isInteger(
          classification.estimatedMinutes
        )
          ? classification.estimatedMinutes
          : null,
      needsReview:
        typeof classification.needsReview === "boolean"
          ? classification.needsReview
          : true,
    };
  }

  function setCaptureButtonBusy(isBusy) {
    if (!captureButton) {
      return;
    }

    captureButton.disabled = isBusy;
    captureButton.textContent = isBusy
      ? "Capturing..."
      : "Capture";
  }

  function analyzeTask(text) {
    const lowerText = text.toLowerCase();

    let category =
      settings.defaultList || "General";

    let priority = "Normal";
    let dueDate = "";
    let context = "";

    if (
      lowerText.includes("supplier") ||
      lowerText.includes("vendor") ||
      lowerText.includes("quote") ||
      lowerText.includes("purchase")
    ) {
      category = "Purchasing";
    } else if (
      lowerText.includes("warehouse") ||
      lowerText.includes("inventory") ||
      lowerText.includes("shipment")
    ) {
      category = "Operations";
    } else if (
      lowerText.includes("team") ||
      lowerText.includes("employee") ||
      lowerText.includes("meeting")
    ) {
      category = "Leadership";
    } else if (
      lowerText.includes("home") ||
      lowerText.includes("family") ||
      lowerText.includes("personal")
    ) {
      category = "Personal";
    } else if (
      lowerText.includes("idea") ||
      lowerText.includes("maybe") ||
      lowerText.includes("someday")
    ) {
      category = "Ideas";
    }

    if (
      lowerText.includes("important") ||
      lowerText.includes("urgent") ||
      lowerText.includes("asap") ||
      lowerText.includes("high priority")
    ) {
      priority = "High";
    }

    if (lowerText.includes("today")) {
      dueDate = getDateString(new Date());
    }

    if (lowerText.includes("tomorrow")) {
      const tomorrow = new Date();
      tomorrow.setDate(tomorrow.getDate() + 1);
      dueDate = getDateString(tomorrow);
    }

    if (
      lowerText.includes("call") ||
      lowerText.includes("phone")
    ) {
      context = "Calls";
    } else if (
      lowerText.includes("computer") ||
      lowerText.includes("email")
    ) {
      context = "Computer";
    } else if (
      lowerText.includes("pick up") ||
      lowerText.includes("store")
    ) {
      context = "Errands";
    } else if (category === "Personal") {
      context = "Home";
    } else {
      context = "Work";
    }

    return {
      summary: text,
      category: category,
      priority: priority,
      dueDate: dueDate,
      context: context,
      project: "",
      estimatedMinutes: null,
      needsReview: true,
    };
  }

  function renderEverything() {
    try {
      renderInbox();
      renderToday();
      renderReview();
      updateCounts();
    } catch (error) {
      console.error(error);

      showCaptureMessage(
        "Display error: " + error.message,
        true
      );
    }
  }

  function renderInbox() {
    const openTasks = tasks.filter(function (task) {
      return !task.completed && !task.localDeleted;
    });

    renderTaskList(
      recentTaskList,
      openTasks,
      "inbox"
    );
  }

  function renderToday() {
    const todayTasks = tasks.filter(function (task) {
      return (
        task.plannedForToday &&
        !task.completed &&
        !task.localDeleted
      );
    });

    renderTaskList(
      todayTaskList,
      todayTasks,
      "today"
    );
  }

  function renderReview() {
    const reviewTasks = tasks.filter(function (task) {
      return (
        task.needsReview &&
        !task.completed &&
        !task.localDeleted
      );
    });

    renderTaskList(
      reviewTaskList,
      reviewTasks,
      "review"
    );
  }

  function renderTaskList(listElement, taskArray, location) {
    listElement.innerHTML = "";

    if (taskArray.length === 0) {
      const emptyBox = document.createElement("div");
      emptyBox.className = "empty-state";

      if (location === "inbox") {
        emptyBox.innerHTML =
          "<span class='empty-icon'>✓</span>" +
          "<h3>Your inbox is clear</h3>" +
          "<p>New tasks and ideas will appear here.</p>";
      }

      if (location === "today") {
        emptyBox.innerHTML =
          "<span class='empty-icon'>☀️</span>" +
          "<h3>No tasks planned yet</h3>" +
          "<p>Add an inbox task to Today.</p>";
      }

      if (location === "review") {
        emptyBox.innerHTML =
          "<span class='empty-icon'>🧠</span>" +
          "<h3>Nothing needs review</h3>" +
          "<p>Only unclear captures will appear here.</p>";
      }

      listElement.appendChild(emptyBox);
      return;
    }

    taskArray.forEach(function (task) {
      const card = document.createElement("article");
      card.className = "task-card";

      const heading = document.createElement("h3");
      heading.textContent = task.title;

      const originalText = document.createElement("p");
      originalText.textContent = task.originalText;

      let suggestionPanel = null;

      if (
        location === "review" &&
        (task.suggestedCategory ||
          Number.isInteger(task.aiConfidence) ||
          task.aiExplanation)
      ) {
        suggestionPanel = document.createElement("div");
        suggestionPanel.className = "review-suggestion";

        const suggestionHeading = document.createElement("strong");
        suggestionHeading.textContent =
          "AI suggests: " +
          (task.suggestedCategory ||
            task.category ||
            "General");
        suggestionPanel.appendChild(suggestionHeading);

        if (Number.isInteger(task.aiConfidence)) {
          const confidenceText = document.createElement("span");
          confidenceText.textContent =
            "Confidence: " + task.aiConfidence + "%";
          suggestionPanel.appendChild(confidenceText);
        }

        if (task.aiExplanation) {
          const explanationText = document.createElement("p");
          explanationText.textContent = task.aiExplanation;
          suggestionPanel.appendChild(explanationText);
        }
      }

      const meta = document.createElement("div");
      meta.className = "task-meta";

      meta.appendChild(
        createTag(task.category || "General")
      );

      meta.appendChild(
        createTag(task.priority || "Normal")
      );

      meta.appendChild(
        createTag(
          task.dueDate
            ? "Due " + formatDate(task.dueDate)
            : "No due date"
        )
      );

      if (task.context) {
        meta.appendChild(
          createTag(task.context)
        );
      }

      if (task.project) {
        meta.appendChild(
          createTag("Project: " + task.project)
        );
      }

      if (task.estimatedMinutes) {
        meta.appendChild(
          createTag(
            task.estimatedMinutes + " min"
          )
        );
      }

      meta.appendChild(
        createTag(getMicrosoftSyncLabel(task))
      );

      const actions = document.createElement("div");
      actions.className = "task-card-actions";

      if (location === "review") {
        actions.appendChild(
          createButton(
            "Approve Suggested Folder",
            "",
            function () {
              approveTaskAndSync(task);
            }
          )
        );

        actions.appendChild(
          createButton(
            "Choose Different Folder",
            "",
            function () {
              chooseDifferentFolderAndApprove(task);
            }
          )
        );

        if (task.serverReviewId) {
          actions.appendChild(
            createButton(
              "Keep in GSD Review",
              "",
              function () {
                keepTaskInGsdReview(task);
              }
            )
          );
        }
      }

      actions.appendChild(
        createButton(
          task.plannedForToday
            ? "Remove from Today"
            : "Add to Today",
          "",
          function () {
            updateTask(task.id, function (item) {
              item.plannedForToday =
                !item.plannedForToday;
            });
          }
        )
      );

      if (
        !task.microsoftTaskId &&
        task.microsoftSyncStatus !== "syncing" &&
        task.microsoftSyncStatus !== "moving"
      ) {
        actions.appendChild(
          createButton(
            "Send to To Do",
            "",
            function () {
              syncTaskToMicrosoft(task);
            }
          )
        );
      }

      actions.appendChild(
        createButton(
          "Complete",
          "complete-button",
          function () {
            completeTaskAndSync(task);
          }
        )
      );

      actions.appendChild(
        createButton(
          "Delete",
          "delete-button",
          function () {
            const shouldDelete = window.confirm(
              'Delete "' + task.title + '"?'
            );

            if (!shouldDelete) {
              return;
            }

            deleteTaskAndSync(task);
          }
        )
      );

      card.appendChild(heading);

      if (task.originalText !== task.title) {
        card.appendChild(originalText);
      }

      if (suggestionPanel) {
        card.appendChild(suggestionPanel);
      }

      card.appendChild(meta);
      card.appendChild(actions);

      listElement.appendChild(card);
    });
  }

  function getMicrosoftSyncLabel(task) {
    if (task.microsoftSyncStatus === "moving") {
      return "To Do: Moving to approved list...";
    }

    if (
      task.needsReview &&
      (task.microsoftReviewTaskId ||
        task.microsoftTaskId ||
        task.microsoftSyncStatus === "synced")
    ) {
      return "To Do: Needs review → " +
        (task.microsoftReviewListName ||
          task.microsoftListName ||
          MICROSOFT_REVIEW_LIST_NAME);
    }

    if (task.microsoftTaskId || task.microsoftSyncStatus === "synced") {
      return task.microsoftListName
        ? "To Do: Synced → " + task.microsoftListName
        : "To Do: Synced";
    }

    if (task.microsoftSyncStatus === "syncing") {
      return task.needsReview
        ? "To Do: Sending to GSD Review..."
        : "To Do: Sending...";
    }

    if (task.microsoftSyncStatus === "failed") {
      return "To Do: Retry needed";
    }

    if (task.needsReview) {
      return "To Do: Review pending locally";
    }

    return "To Do: Local only";
  }

  function createTag(text) {
    const tag = document.createElement("span");
    tag.className = "task-tag";
    tag.textContent = text;
    return tag;
  }

  function createButton(text, extraClass, clickHandler) {
    const button = document.createElement("button");

    button.type = "button";
    button.className =
      "small-action-button " + extraClass;
    button.textContent = text;
    button.addEventListener("click", clickHandler);

    return button;
  }

  function updateTask(taskId, updateFunction) {
    const task = tasks.find(function (item) {
      return item.id === taskId;
    });

    if (!task) {
      return;
    }

    updateFunction(task);
    saveTasks();
    renderEverything();
  }

  function updateCounts() {
    const visibleTasks = tasks.filter(function (task) {
      return !task.localDeleted;
    });

    const openTasks = visibleTasks.filter(function (task) {
      return !task.completed;
    });

    const completedTasks = visibleTasks.filter(function (task) {
      return task.completed;
    });

    const reviewTasks = visibleTasks.filter(function (task) {
      return task.needsReview && !task.completed;
    });

    if (inboxCount) {
      inboxCount.textContent = openTasks.length;
    }

    if (reviewCount) {
      reviewCount.textContent = reviewTasks.length;
    }

    if (capturedTotal) {
      capturedTotal.textContent = visibleTasks.length;
    }

    if (completedTotal) {
      completedTotal.textContent =
        completedTasks.length;
    }

    if (needsReviewTotal) {
      needsReviewTotal.textContent =
        reviewTasks.length;
    }
  }

  function showScreen(screenId) {
    screens.forEach(function (screen) {
      if (screen.id === screenId) {
        screen.classList.add("active-screen");
      } else {
        screen.classList.remove("active-screen");
      }
    });

    navButtons.forEach(function (button) {
      if (button.dataset.screen === screenId) {
        button.classList.add("active-nav");
      } else {
        button.classList.remove("active-nav");
      }
    });

    window.scrollTo(0, 0);
  }

  function saveSettings() {
    settings = {
      userName: userNameInput
        ? userNameInput.value.trim()
        : "",
      defaultList:
        defaultListInput &&
        defaultListInput.value.trim()
          ? defaultListInput.value.trim()
          : "General",
      aiSorting: aiSortingToggle
        ? aiSortingToggle.checked
        : true,
      calendarSuggestions: calendarToggle
        ? calendarToggle.checked
        : true,
      aiListRouting: microsoftRoutingModeSelect
        ? microsoftRoutingModeSelect.value
        : settings.aiListRouting || "ask",
      microsoftListId: settings.microsoftListId || "",
      microsoftListName: settings.microsoftListName || "",
    };

    localStorage.setItem(
      SETTINGS_KEY,
      JSON.stringify(settings)
    );

    if (settingsMessage) {
      settingsMessage.textContent =
        "Settings saved.";
    }
  }

  function populateSettings() {
    if (userNameInput) {
      userNameInput.value =
        settings.userName || "";
    }

    if (defaultListInput) {
      defaultListInput.value =
        settings.defaultList || "General";
    }

    if (aiSortingToggle) {
      aiSortingToggle.checked =
        settings.aiSorting !== false;
    }

    if (calendarToggle) {
      calendarToggle.checked =
        settings.calendarSuggestions !== false;
    }

    if (microsoftRoutingModeSelect) {
      microsoftRoutingModeSelect.value =
        settings.aiListRouting || "ask";
    }

    updateMicrosoftRoutingMessage();
  }

  function startVoiceCapture() {
    const SpeechRecognition =
      window.SpeechRecognition ||
      window.webkitSpeechRecognition;

    if (!SpeechRecognition) {
      showCaptureMessage(
        "Use the microphone on your keyboard to dictate the task.",
        true
      );
      return;
    }

    const recognition = new SpeechRecognition();

    recognition.lang = "en-US";
    recognition.interimResults = false;

    voiceButton.disabled = true;
    voiceButton.textContent = "Listening...";

    recognition.start();

    recognition.onresult = function (event) {
      taskInput.value =
        event.results[0][0].transcript;
    };

    recognition.onerror = function () {
      showCaptureMessage(
        "Voice capture did not complete.",
        true
      );
    };

    recognition.onend = function () {
      voiceButton.disabled = false;
      voiceButton.textContent = "🎙️ Speak";
    };
  }

  function saveTasks() {
    localStorage.setItem(
      TASKS_KEY,
      JSON.stringify(tasks)
    );
  }

  function loadTasks() {
    try {
      const saved = localStorage.getItem(TASKS_KEY);

      if (!saved) {
        return [];
      }

      const parsed = JSON.parse(saved);

      return Array.isArray(parsed) ? parsed : [];
    } catch (error) {
      console.error(error);
      return [];
    }
  }

  function loadSettings() {
    try {
      const saved = localStorage.getItem(SETTINGS_KEY);

      if (!saved) {
        return {
          userName: "",
          defaultList: "General",
          aiSorting: true,
          calendarSuggestions: true,
          aiListRouting: "ask",
          microsoftListId: "",
          microsoftListName: "",
        };
      }

      const parsed = JSON.parse(saved);

      return {
        userName: "",
        defaultList: "General",
        aiSorting: true,
        calendarSuggestions: true,
        aiListRouting: "ask",
        microsoftListId: "",
        microsoftListName: "",
        ...(parsed && typeof parsed === "object" ? parsed : {}),
      };
    } catch (error) {
      return {
        userName: "",
        defaultList: "General",
        aiSorting: true,
        calendarSuggestions: true,
        aiListRouting: "ask",
        microsoftListId: "",
        microsoftListName: "",
      };
    }
  }

  function showCaptureMessage(message, isError) {
    if (!captureMessage) {
      window.alert(message);
      return;
    }

    captureMessage.textContent = message;
    captureMessage.style.color = isError
      ? "#c73535"
      : "#137a52";
  }

  function getDateString(date) {
    const year = date.getFullYear();
    const month = String(
      date.getMonth() + 1
    ).padStart(2, "0");
    const day = String(date.getDate()).padStart(
      2,
      "0"
    );

    return year + "-" + month + "-" + day;
  }

  function formatDate(dateString) {
    const date = new Date(
      dateString + "T12:00:00"
    );

    return date.toLocaleDateString("en-US", {
      month: "short",
      day: "numeric",
      year: "numeric",
    });
  }
});

if ("serviceWorker" in navigator) {
  window.addEventListener("load", function () {
    navigator.serviceWorker
      .register("/sw.js")
      .then(function () {
        console.log("GSD Capture service worker registered.");
      })
      .catch(function (error) {
        console.error("Service worker registration failed:", error);
      });
  });
}
