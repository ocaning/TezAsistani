/**
 * Word için Gemini Tabanlı Akademik Tez Asistanı - taskpane.js
 */

const GEMINI_API_ENDPOINT = "https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent";
const SECURITY_SALT = "BeunMed_TezAsistani_2026_X9"; 

const SYSTEM_INSTRUCTION = `Sen tıp doktorları ve akademisyenler için uzman bir Akademik Tez Asistanı ve Tıbbi Editörsün.
Kritik Kurallar:
1. Tıbbi terminolojideki hassasiyetlere kesinlikle uy (Örn: inflamasyon, inflamatuvar kelimelerini mutlaka 'v' ile yaz; 'Round Ligament' gibi anatomik yapıları evrensel cerrahi standardına göre etiketle).
2. Word dokümanından gelen metnin hipotez sunumunu ve istatistiksel dilini tıp literatürüne (pasif akademik dile) uygun hale getir.
3. Kanıta dayalı tıp ilkelerini koru, spekülasyondan kaçın.`;

// --- Şifreleme Fonksiyonları ---
function encryptKey(plainText) {
    let encrypted = "";
    for (let i = 0; i < plainText.length; i++) {
        encrypted += String.fromCharCode(plainText.charCodeAt(i) ^ SECURITY_SALT.charCodeAt(i % SECURITY_SALT.length));
    }
    return btoa(encrypted); 
}

function decryptKey(encodedText) {
    try {
        let decoded = atob(encodedText);
        let decrypted = "";
        for (let i = 0; i < decoded.length; i++) {
            decrypted += String.fromCharCode(decoded.charCodeAt(i) ^ SECURITY_SALT.charCodeAt(i % SECURITY_SALT.length));
        }
        return decrypted;
    } catch (e) { return null; }
}

// --- Ana Uygulama Mantığı ---
Office.onReady((info) => {
    // UI Elemanlarını Tanımla
    const loginScreen = document.getElementById("login-screen");
    const chatScreen = document.getElementById("chat-screen");
    const apiKeyInput = document.getElementById("api-key-input");
    const btnLogin = document.getElementById("btn-login");
    const btnLogout = document.getElementById("btn-logout");
    const btnSend = document.getElementById("btn-send");
    const chatInput = document.getElementById("chat-input");
    const chatMessages = document.getElementById("chat-messages");
    const textArea = document.getElementById("myTextArea");

    // 1. WebView2 (C#) Köprüsünü Dinle
    window.chrome.webview.addEventListener('message', event => {
        const data = event.data;
        if (data && data.type === 'wordContent') {
            if (textArea) {
                textArea.value = data.text;
                console.log("Metin başarıyla aktarıldı.");
            }
        }
    });

    // 2. Buton Olaylarını Bağla
    if (btnLogin) btnLogin.addEventListener("click", handleLogin);
    if (btnLogout) btnLogout.addEventListener("click", handleLogout);
    if (btnSend) btnSend.addEventListener("click", handleSendMessage);
    if (chatInput) {
        chatInput.addEventListener("keypress", (e) => {
            if (e.key === "Enter" && !e.shiftKey) { e.preventDefault(); handleSendMessage(); }
        });
    }

    // 3. Başlatma
    if (info.host === Office.HostType.Word) {
        checkAuth();
    }

    // --- Fonksiyonlar ---
    function checkAuth() {
        const encryptedKey = localStorage.getItem("tez_asistani_api_key_secure");
        if (encryptedKey && encryptedKey.trim() !== "") {
            const decryptedKey = decryptKey(encryptedKey);
            if (decryptedKey && decryptedKey.startsWith("AIza")) {
                loginScreen.style.display = "none";
                chatScreen.style.display = "flex";
                if(chatMessages.innerHTML === "") appendSystemMessage("Sistem hazır. Word'den metin seçerek soru sorabilirsiniz.");
                return;
            }
        }
        loginScreen.style.display = "flex";
        chatScreen.style.display = "none";
    }

    function handleLogin() {
        const key = apiKeyInput.value.trim();
        if (key === "" || !key.startsWith("AIza")) {
            alert("Lütfen AIza... ile başlayan geçerli bir Google AI Studio anahtarı giriniz.");
            return;
        }
        localStorage.setItem("tez_asistani_api_key_secure", encryptKey(key));
        checkAuth();
    }

    function handleLogout() {
        if (confirm("Oturumu kapatırsanız API anahtarınız bu bilgisayardan silinir. Emin misiniz?")) {
            localStorage.removeItem("tez_asistani_api_key_secure");
            chatMessages.innerHTML = "";
            checkAuth();
        }
    }

    async function getSelectedText() {
        return new Promise((resolve) => {
            Word.run(async (context) => {
                const selection = context.document.getSelection();
                selection.load("text");
                await context.sync();
                resolve(selection.text && selection.text.trim() !== "" ? selection.text.trim() : null);
            }).catch(() => resolve(null));
        });
    }

    async function handleSendMessage() {
        const query = chatInput.value.trim();
        if (query === "") return;

        const apiKey = decryptKey(localStorage.getItem("tez_asistani_api_key_secure"));
        if (!apiKey) { checkAuth(); return; }

        appendUserMessage(query);
        chatInput.value = "";
        const loadingIndicator = appendLoadingMessage();

        try {
            // Öncelik sırası: myTextArea doluysa onu al, yoksa Word'den seçili metni al
            const selectedText = (textArea && textArea.value.trim() !== "") ? textArea.value : await getSelectedText();
            
            let promptContent = selectedText 
                ? `[Seçili Metin]:\n"${selectedText}"\n\n[Kullanıcı Talimatı]:\n${query}` 
                : `[Kullanıcı Talimatı]:\n${query}`;

            const response = await fetch(`${GEMINI_API_ENDPOINT}?key=${apiKey}`, {
                method: "POST",
                headers: { "Content-Type": "application/json" },
                body: JSON.stringify({
                    contents: [{ role: "user", parts: [{ text: promptContent }] }],
                    systemInstruction: { parts: [{ text: SYSTEM_INSTRUCTION }] },
                    generationConfig: { temperature: 0.2, topP: 0.95 }
                })
            });

            removeMessageElement(loadingIndicator);
            if (!response.ok) throw new Error("Ağ yanıtı başarısız.");
            
            const data = await response.json();
            const aiResponse = data.candidates?.[0]?.content?.parts?.[0]?.text;
            if (aiResponse) appendAiMessage(aiResponse);

        } catch (error) {
            removeMessageElement(loadingIndicator);
            appendErrorMessage("Hata: İşlem gerçekleştirilemedi.");
        }
    }

    // Yardımcı UI Fonksiyonları
    function appendUserMessage(text) { const msg = document.createElement("div"); msg.className = "message user-message"; msg.innerText = text; chatMessages.appendChild(msg); scrollToBottom(); }
    function appendAiMessage(text) { const msg = document.createElement("div"); msg.className = "message ai-message"; msg.innerText = text; chatMessages.appendChild(msg); scrollToBottom(); }
    function appendSystemMessage(text) { const msg = document.createElement("div"); msg.className = "message system-message"; msg.innerText = text; chatMessages.appendChild(msg); scrollToBottom(); }
    function appendErrorMessage(text) { const msg = document.createElement("div"); msg.className = "message error-message"; msg.innerText = text; chatMessages.appendChild(msg); scrollToBottom(); }
    function appendLoadingMessage() { const msg = document.createElement("div"); msg.className = "message ai-message system-message"; msg.innerText = "Akademik yanıt üretiliyor..."; chatMessages.appendChild(msg); scrollToBottom(); return msg; }
    function removeMessageElement(element) { if (element && element.parentNode) element.parentNode.removeChild(element); }
    function scrollToBottom() { chatMessages.scrollTop = chatMessages.scrollHeight; }
});