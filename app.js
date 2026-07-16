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

  const captureForm = document.getElementById("captureForm");
  const taskInput = document.getElementById("taskInput");
  const captureMessage = document.getElementById("captureMessage");
  const recentTaskList = document.getElementById("recentTaskList");
  const todayTaskList = document.getElementById("todayTaskList");
  const reviewTaskList = document.getElementById("reviewTaskList");

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

  function registerEvents() {
    captureForm.addEventListener("submit", captureTask);

    navButtons.forEach(function (button) {
      button.addEventListener("click", function () {
        showScreen(button.dataset.screen);
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
      "Approved captures will sync automatically to " +
        settings.microsoftListName +
        ". You can still use the test button to verify the connection."
    );
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

    if (task.needsReview) {
      task.microsoftSyncStatus = "waiting-review";
      saveTasks();
      renderEverything();
      return { status: "waiting-review" };
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
    task.microsoftListId = getMicrosoftListIdForTask(task);
    task.microsoftListName = getMicrosoftListNameForTask(task);
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

  async function createMicrosoftTask(task, accessToken) {
    const controller = new AbortController();
    const timeoutId = window.setTimeout(function () {
      controller.abort();
    }, 15000);

    let response;

    try {
      response = await fetch(
        "https://graph.microsoft.com/v1.0/me/todo/lists/" +
          encodeURIComponent(getMicrosoftListIdForTask(task)) +
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

  async function findExistingMicrosoftTask(task, accessToken) {
    let nextUrl =
      "https://graph.microsoft.com/v1.0/me/todo/lists/" +
      encodeURIComponent(getMicrosoftListIdForTask(task)) +
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
    task.microsoftSyncStatus = "synced";
    task.microsoftSyncError = "";
    task.microsoftSyncedAt = new Date().toISOString();
    saveTasks();
    renderEverything();
  }

  async function approveTaskAndSync(task) {
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

    await syncTaskToMicrosoft(task);
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
        microsoftSyncStatus:
          typeof analysis.needsReview === "boolean" &&
          analysis.needsReview === false
            ? "pending"
            : "waiting-review",
        microsoftSyncError: "",
        microsoftSyncAttemptedAt: "",
        microsoftSyncedAt: "",
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

      if (newTask.needsReview) {
        showCaptureMessage(
          baseMessage +
            " Approve it before sending it to Microsoft To Do."
        );
        return;
      }

      if (!canSyncTaskToMicrosoft(newTask)) {
        newTask.microsoftSyncStatus = "local-only";
        saveTasks();
        renderEverything();

        showCaptureMessage(
          baseMessage +
            " Saved locally. Connect Microsoft To Do and select a list to send it later."
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
            " Added to " +
            (newTask.microsoftListName ||
              settings.microsoftListName ||
              "Microsoft To Do") +
            "."
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
      return !task.completed;
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
        !task.completed
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
        !task.completed
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
            "Approve",
            "",
            function () {
              approveTaskAndSync(task);
            }
          )
        );
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
        !task.needsReview &&
        !task.microsoftTaskId &&
        task.microsoftSyncStatus !== "syncing"
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
            updateTask(task.id, function (item) {
              item.completed = true;
              item.needsReview = false;
              item.plannedForToday = false;
            });
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

            tasks = tasks.filter(function (item) {
              return item.id !== task.id;
            });

            saveTasks();
            renderEverything();
          }
        )
      );

      card.appendChild(heading);

      if (task.originalText !== task.title) {
        card.appendChild(originalText);
      }

      card.appendChild(meta);
      card.appendChild(actions);

      listElement.appendChild(card);
    });
  }

  function getMicrosoftSyncLabel(task) {
    if (task.microsoftTaskId || task.microsoftSyncStatus === "synced") {
      return "To Do: Synced";
    }

    if (task.needsReview || task.microsoftSyncStatus === "waiting-review") {
      return "To Do: Waiting for review";
    }

    if (task.microsoftSyncStatus === "syncing") {
      return "To Do: Sending...";
    }

    if (task.microsoftSyncStatus === "failed") {
      return "To Do: Retry needed";
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
    const openTasks = tasks.filter(function (task) {
      return !task.completed;
    });

    const completedTasks = tasks.filter(function (task) {
      return task.completed;
    });

    const reviewTasks = tasks.filter(function (task) {
      return task.needsReview && !task.completed;
    });

    if (inboxCount) {
      inboxCount.textContent = openTasks.length;
    }

    if (reviewCount) {
      reviewCount.textContent = reviewTasks.length;
    }

    if (capturedTotal) {
      capturedTotal.textContent = tasks.length;
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
