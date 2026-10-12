// Handle uploads, chat streaming, history, rendering, and user interactions.

let currentMessages = [];
let currentConversationId = null;
let activeStreamState = null;

function getMessageAvatar(role) {
    return role === "user" ? "You" : "AI";
}

function createConversationId() {
    return typeof crypto !== "undefined" && crypto.randomUUID
        ? crypto.randomUUID()
        : `conversation-${Date.now()}-${Math.random().toString(16).slice(2)}`;
}

function ensureConversationId() {
    if (!currentConversationId) {
        currentConversationId = createConversationId();
    }
    return currentConversationId;
}

async function startNewChat() {
    currentMessages = [];
    currentConversationId = null;

    const chatWindow = document.getElementById("chat-window");
    if (chatWindow) {
        chatWindow.innerHTML = `
            <div class="message ai-message-row">
                <div class="message-header">
                    <span class="avatar-badge ai-avatar">AI</span>
                </div>
                <div class="ai-message">
                    <h3>Welcome to COMM_IT AI</h3>
                    <p>Upload files or ask questions.</p>
                </div>
                <div class="message-time">
                    ${formatMessageTime()}
                </div>
            </div>
        `;
    }

    const messageBox = document.getElementById("message");
    if (messageBox) {
        messageBox.value = "";
        messageBox.focus();
    }

    const fileInput = document.getElementById("fileUpload");
    if (fileInput) {
        fileInput.value = "";
    }

    renderSelectedFiles();

    try {
        await fetch("/new-chat", { method: "POST" });
    }
    catch (error) {
        console.error("Failed to reset server chat state:", error);
    }
}

function formatMessageTime(date = new Date()) {

    return date.toLocaleTimeString(
        [],
        {
            hour: "2-digit",
            minute: "2-digit"
        }
    );
}

function removeSourceCitations(content) {

    return content.replace(
        /\s*\[Source:\s*[^\]]+\]/gi,
        ""
    ).trim();
}

function triggerDownloadUrl(url) {
    if (!url) {
        return;
    }

    const anchor = document.createElement("a");
    anchor.href = url;
    anchor.rel = "noopener noreferrer";

    if (url.startsWith("/download?")) {
        const filenameMatch = url.match(/[?&]filename=([^&]+)/);
        if (filenameMatch) {
            anchor.download = decodeURIComponent(filenameMatch[1]);
        }
    } else if (
        !url.startsWith("/download/tool/") &&
        !url.startsWith("/download/maximo/")
    ) {
        const fileName = url.split("/").pop() || "download.pdf";
        anchor.download = fileName;
    }

    document.body.appendChild(anchor);
    anchor.click();
    anchor.remove();
}

/* =====================================
   SELECTED FILES
===================================== */

function renderSelectedFiles() {

    const fileInput =
        document.getElementById(
            "fileUpload"
        );

    const selectedFiles =
        document.getElementById(
            "selected-files"
        );

    if (!fileInput || !selectedFiles) {
        return;
    }

    selectedFiles.innerHTML = "";

    Array.from(fileInput.files).forEach(
        (file, index) => {

            const fileItem =
                document.createElement(
                    "div"
                );

            fileItem.className =
                "selected-file";

            fileItem.innerHTML = `
                <span class="selected-file-name" title="${escapeHtml(file.name)}">
                    ${escapeHtml(file.name)}
                </span>
                <button
                    type="button"
                    class="remove-file-button"
                    aria-label="Remove ${escapeHtml(file.name)}"
                    title="Remove file"
                >
                    ×
                </button>
            `;

            fileItem
                .querySelector(
                    ".remove-file-button"
                )
                .addEventListener(
                    "click",
                    () => removeSelectedFile(index)
                );

            selectedFiles.appendChild(
                fileItem
            );
        }
    );
}


function removeSelectedFile(fileIndex) {

    const fileInput =
        document.getElementById(
            "fileUpload"
        );

    if (!fileInput) {
        return;
    }

    const remainingFiles =
        Array.from(fileInput.files).filter(
            (_, index) => index !== fileIndex
        );

    const dataTransfer =
        new DataTransfer();

    remainingFiles.forEach(
        (file) => dataTransfer.items.add(file)
    );

    fileInput.files = dataTransfer.files;
    renderSelectedFiles();
}


document
    .getElementById("fileUpload")
    ?.addEventListener(
        "change",
        renderSelectedFiles
    );

/* =====================================
   FILE UPLOAD
===================================== */

async function uploadFiles() {

    const fileInput =
        document.getElementById(
            "fileUpload"
        );

    if (!fileInput) {
        return;
    }

    const files =
        Array.from(fileInput.files);

    if (files.length === 0) {
        return;
    }

    for (const [index, file] of files.entries()) {

        const formData =
            new FormData();

        formData.append(
            "file",
            file
        );

        const response =
            await fetch(
                `/upload?replace_session=${index === 0}`,
                {
                    method: "POST",
                    body: formData
                }
            );

        const responseText =
            await response.text();
        let result = {};

        try {
            result = responseText
                ? JSON.parse(responseText)
                : {};
        }
        catch {
            result = {
                error: responseText || response.statusText
            };
        }

        if (!response.ok || !result.success) {
            throw new Error(
                result.error ||
                `Upload failed for ${file.name}`
            );
        }

        console.log(
            "Uploaded:",
            result
        );

    }

    fileInput.value = "";
    renderSelectedFiles();

}

/* =====================================
   SEND MESSAGE
===================================== */

async function sendMessage() {

    if (activeStreamState) {
        return;
    }

    const messageBox =
        document.getElementById(
            "message"
        );

    const message =
        messageBox.value.trim();

    if (!message) {
        return;
    }

    try {
        await uploadFiles();
    }
    catch (error) {
        console.error(error);

        const chatWindow =
            document.getElementById(
                "chat-window"
            );

        chatWindow.innerHTML += `
            <div class="message">
                <div class="ai-message">
                    Upload error: ${escapeHtml(error.message)}
                </div>
            </div>
        `;

        chatWindow.scrollTop =
            chatWindow.scrollHeight;
        return;
    }

    const chatWindow =
        document.getElementById(
            "chat-window"
        );

    chatWindow.innerHTML += `
        <div class="message user-message-row">
            <div class="message-header user-header">
                <span class="avatar-badge user-avatar">You</span>
            </div>
            <div class="user-message">
                ${escapeHtml(message)}
            </div>
            <div class="message-time">
                ${formatMessageTime()}
            </div>
        </div>
    `;

    currentMessages.push({
        role: "user",
        content: message,
        timestamp: new Date().toISOString()
    });

    messageBox.value = "";

    const loadingId =
        "loading-" + Date.now();
    const streamState = {
        paused: false,
        reader: null,
        answer: "",
        abortController: new AbortController()
    };
    activeStreamState = streamState;
    const pauseButton = document.getElementById("pauseStreamButton");
    const sendButton = document.getElementById("sendButton");
    pauseButton.hidden = false;
    sendButton.disabled = true;

    chatWindow.innerHTML += `
        <div
            class="message ai-message-row"
            id="${loadingId}"
        >
            <div class="message-header">
                <span class="avatar-badge ai-avatar">AI</span>
            </div>
            <div class="ai-message">
                <div class="streaming-body">Thinking...</div>
            </div>
            <div class="message-time">
                ${formatMessageTime()}
            </div>
        </div>
    `;

    const loadingElement = document.getElementById(loadingId);
    const handlePause = () => {
        if (!streamState.paused) {
            streamState.paused = true;
            pauseButton.hidden = true;
            if (streamState.reader) {
                streamState.reader.cancel().catch(() => {});
            } else {
                streamState.abortController.abort();
            }
        }
    };
    pauseButton.addEventListener("click", handlePause);

    chatWindow.scrollTop =
        chatWindow.scrollHeight;

    try {

        const response =
            await fetch(
                "/chat/stream",
                {
                    method: "POST",
                    headers: {
                        "Content-Type":
                            "application/json"
                    },
                    signal: streamState.abortController.signal,
                    body: JSON.stringify({
                        message: message,
                        history: currentMessages.slice(0, -1).slice(-8)
                    })
                }
            );

        if (!response.ok) {

            throw new Error(
                `HTTP ${response.status}`
            );

        }

        const reader =
            response.body.getReader();
        streamState.reader = reader;

        const decoder =
            new TextDecoder();

        let answer = "";
        let downloadTriggered = false;

        const updateAnswer = (isFinal = false) => {

            const visibleAnswer =
                removeSourceCitations(answer);

            const answerElement =
                loadingElement?.querySelector(
                    ".streaming-body"
                );

            const trimmedAnswer = visibleAnswer.trim();
            if (
                trimmedAnswer.startsWith("/download?url=") ||
                trimmedAnswer.startsWith("/download/tool/") ||
                trimmedAnswer.startsWith("/download/maximo/")
            ) {
                if (answerElement) {
                    answerElement.textContent = isFinal
                        ? "Download started. Check your browser's downloads."
                        : "Preparing download…";
                }
                if (isFinal && !downloadTriggered) {
                    downloadTriggered = true;
                    triggerDownloadUrl(trimmedAnswer);
                }
                return;
            }

            if (/^https?:\/\//i.test(trimmedAnswer)) {
                if (answerElement) {
                    answerElement.innerHTML = "<span>File download is only available for valid PM procedure requests.</span>";
                }
                return;
            }

            let renderedHtml =
                escapeHtml(visibleAnswer);

            if (
                typeof marked !==
                "undefined"
            ) {

                renderedHtml =
                    marked.parse(visibleAnswer);
                renderedHtml =
                    renderedHtml.replace(
                        /<a\s+/gi,
                        '<a target="_blank" rel="noopener noreferrer" '
                    );
            }

            if (answerElement) {
                answerElement.innerHTML =
                    renderedHtml;
            }
        };

        while (true) {
            if (streamState.paused) {
                break;
            }

            const result = await reader.read();

            answer += decoder.decode(
                result.value || new Uint8Array(),
                {
                    stream: !result.done
                }
            );

            if (answer) {
                updateAnswer();
                chatWindow.scrollTop =
                    chatWindow.scrollHeight;
            }

            if (result.done) {
                break;
            }
        }

        updateAnswer(true);
        if (streamState.paused && !answer) {
            const body = loadingElement?.querySelector(".streaming-body");
            if (body) {
                body.textContent = "Response paused.";
            }
        }
        pauseButton.hidden = true;

        currentMessages.push({
            role: "assistant",
            content: removeSourceCitations(answer).startsWith("/download/")
                ? "Download started. Check your browser's downloads."
                : removeSourceCitations(answer),
            timestamp: new Date().toISOString()
        });

        if (
            typeof hljs !==
            "undefined"
        ) {

            highlightCode();

        }

        addCopyButtons();

        await saveChat();

    }
    catch (error) {

        if (streamState.paused) {
            const pausedElement =
                document.getElementById(
                    loadingId
                );

            if (pausedElement) {
                const body = pausedElement.querySelector(".streaming-body");
                if (body && body.textContent.trim() === "Thinking...") {
                    body.textContent = "Response paused.";
                }
            }

            chatWindow.scrollTop =
                chatWindow.scrollHeight;
            return;
        }

        console.error(error);

        const loadingElement =
            document.getElementById(
                loadingId
            );

        if (loadingElement) {
            loadingElement.remove();
        }

        chatWindow.innerHTML += `
            <div class="message">
                <div class="ai-message">
                    Error:
                    ${error.message}
                </div>
            </div>
        `;
    }

    finally {
        pauseButton.hidden = true;
        pauseButton.removeEventListener("click", handlePause);
        sendButton.disabled = false;
        if (activeStreamState === streamState) {
            activeStreamState = null;
        }
    }

    chatWindow.scrollTop =
        chatWindow.scrollHeight;

}

/* =====================================
   SAVE CHAT
===================================== */

async function saveChat() {

    if (
        currentMessages.length === 0
    ) {
        return;
    }

    try {

        ensureConversationId();

        await fetch(
            "/save-chat",
            {
                method: "POST",
                headers: {
                    "Content-Type":
                        "application/json"
                },
                body: JSON.stringify({

                    title:
                        currentMessages[0]
                        .content
                        .substring(0, 40),

                    messages:
                        currentMessages,
                    conversation_id:
                        currentConversationId

                })
            }
        );

        loadHistory();

    }
    catch (error) {

        console.error(
            error
        );

    }

}

/* =====================================
   LOAD HISTORY
===================================== */

async function loadHistory() {

    try {

        const response =
            await fetch(
                "/history"
            );

        const chats =
            await response.json();

        const container =
            document.getElementById(
                "conversation-list"
            );

        if (!container) {
            return;
        }

        container.innerHTML = "";

        chats.forEach(chat => {

            const row = document.createElement("div");
            row.className = "chat-row";

            const item = document.createElement("div");
            item.className = "chat-item";
            item.textContent = chat;
            item.dataset.filename = chat;
            item.addEventListener("click", () => loadChat(chat));

            const deleteButton = document.createElement("button");
            deleteButton.type = "button";
            deleteButton.className = "delete-btn";
            deleteButton.textContent = "🗑";
            deleteButton.setAttribute("aria-label", `Delete ${chat}`);
            deleteButton.addEventListener("click", (event) => {
                event.stopPropagation();
                deleteChat(chat);
            });

            row.appendChild(item);
            row.appendChild(deleteButton);
            container.appendChild(row);

        });

    }
    catch (error) {

        console.error(
            error
        );

    }

}

/* =====================================
   LOAD CHAT
===================================== */

async function loadChat(
    filename
) {

    try {

        const response =
            await fetch(
                `/history/${encodeURIComponent(filename)}`
            );

        const messages =
            await response.json();

        currentMessages =
            [...messages];

        const chatWindow =
            document.getElementById(
                "chat-window"
            );

        chatWindow.innerHTML = "";

        messages.forEach(msg => {

            let content =
                escapeHtml(
                    removeSourceCitations(msg.content)
                );

            if (
                msg.role ===
                "assistant"
            ) {

                if (
                    typeof marked !==
                    "undefined"
                ) {

                    content =
                        marked.parse(
                                removeSourceCitations(msg.content)
                        );

                }

            }

            chatWindow.innerHTML += `
                <div class="message ${msg.role === "user" ? "user-message-row" : "ai-message-row"}">
                    <div class="message-header ${msg.role === "user" ? "user-header" : ""}">
                        <span class="avatar-badge ${msg.role === "user" ? "user-avatar" : "ai-avatar"}">${getMessageAvatar(msg.role)}</span>
                    </div>

                    <div class="${
                        msg.role === "user"
                        ? "user-message"
                        : "ai-message"
                    }">

                        ${content}

                    </div>

                    ${msg.timestamp ? `
                        <div class="message-time">
                            ${formatMessageTime(new Date(msg.timestamp))}
                        </div>
                    ` : ""}

                </div>
            `;

        });

        if (
            typeof hljs !==
            "undefined"
        ) {

            highlightCode();

        }

        addCopyButtons();

    }
    catch (error) {

        console.error(
            error
        );

    }

}

/* =====================================
   DELETE CHAT
===================================== */

async function deleteChat(
    filename
) {

    const confirmed =
        confirm(
            `Delete ${filename}?`
        );

    if (!confirmed) {
        return;
    }

    await fetch(
        `/history/${encodeURIComponent(filename)}`,
        {
            method: "DELETE"
        }
    );

    loadHistory();

}

/* =====================================
   CODE HIGHLIGHT
===================================== */

function highlightCode() {

    document
        .querySelectorAll(
            "pre code"
        )
        .forEach(block => {

            hljs.highlightElement(
                block
            );

        });

}

/* =====================================
   COPY BUTTONS
===================================== */

function addCopyButtons() {

    document
        .querySelectorAll(
            "pre"
        )
        .forEach(block => {

            if (
                block.querySelector(
                    ".copy-btn"
                )
            ) {
                return;
            }

            const button =
                document.createElement(
                    "button"
                );

            button.className =
                "copy-btn";

            button.innerText =
                "Copy";

            button.onclick =
                async () => {

                    const code =
                        block.querySelector(
                            "code"
                        );

                    if (!code) {
                        return;
                    }

                    await navigator
                        .clipboard
                        .writeText(
                            code.innerText
                        );

                    button.innerText =
                        "Copied!";

                    setTimeout(
                        () => {

                            button.innerText =
                                "Copy";

                        },
                        2000
                    );

                };

            block.appendChild(
                button
            );

        });

}

/* =====================================
   HELPERS
===================================== */

function escapeHtml(text) {

    const div =
        document.createElement(
            "div"
        );

    div.innerText = text;

    return div.innerHTML;

}

/* =====================================
   STARTUP
===================================== */

document.addEventListener(
    "DOMContentLoaded",
    () => {

        loadHistory();

        const messageBox =
            document.getElementById(
                "message"
            );

        document
            .getElementById("newChatButton")
            ?.addEventListener(
                "click",
                startNewChat
            );

        if (
            messageBox
        ) {

            messageBox.addEventListener(
                "keydown",
                function(event) {

                    if (
                        event.key === "Enter" &&
                        !event.shiftKey
                    ) {

                        event.preventDefault();

                        sendMessage();

                    }

                }
            );

        }

        console.log(
            "COMM_IT AI Ready"
        );

    }
);