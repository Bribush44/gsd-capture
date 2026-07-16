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

    setMicrosoftMessage(
      "Sign in to connect GSD Capture to your Microsoft account."
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
      };

      tasks.unshift(newTask);
      saveTasks();

      taskInput.value = "";
      renderEverything();

      if (usedAi) {
        showCaptureMessage(
          newTask.needsReview
            ? "Captured and sorted with AI. This item needs review."
            : "Captured and sorted with AI."
        );
      } else {
        showCaptureMessage(
          navigator.onLine
            ? "Captured successfully using local sorting."
            : "Captured offline using local sorting."
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

      const actions = document.createElement("div");
      actions.className = "task-card-actions";

      if (location === "review") {
        actions.appendChild(
          createButton(
            "Approve",
            "",
            function () {
              updateTask(task.id, function (item) {
                item.needsReview = false;
              });
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
        };
      }

      return JSON.parse(saved);
    } catch (error) {
      return {
        userName: "",
        defaultList: "General",
        aiSorting: true,
        calendarSuggestions: true,
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
