/* global Office */

Office.onReady((info) => {
    if (info.host === Office.HostType.PowerPoint) {
        document.getElementById("btnGenerateIcon").onclick = handleTextToIcon;
        document.getElementById("btnInsert").onclick = insertToSlide;
    }
});

async function handleTextToIcon() {
    // এলিমেন্টগুলো সংগ্রহ করা
    const iconInput = document.getElementById("iconInput");
    const apiType = document.getElementById("apiType");
    const userApiKey = document.getElementById("userApiKey");
    const loadingDiv = document.getElementById("loading");
    const resultArea = document.getElementById("resultArea");

    // ভ্যালু চেক করা
    const rawInput = iconInput ? iconInput.value : "";
    const selectedApi = apiType ? apiType.value : "huggingface";
    const apiKey = userApiKey ? userApiKey.value : "";
    const styleElement = document.querySelector('input[name="iconStyle"]:checked');
    const style = styleElement ? styleElement.value : "napkin";

    if (!rawInput) {
        alert("Please describe your icon first.");
        return;
    }

    // ১. ব্লিংকিং এবং লোডিং শুরু করা
    if (loadingDiv) loadingDiv.style.display = "block";
    if (resultArea) resultArea.style.display = "none";

    try {
        const serverUrl = "https://suvajit01-sr-visuals-backend.hf.space/generate-icon";
        
        const response = await fetch(serverUrl, {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({ 
                prompt: rawInput,
                style: style,
                api_type: selectedApi,
                user_key: apiKey 
            }),
        });

        const data = await response.json();

        if (data.svg) {
            // ২. প্রিভিউ কন্টেইনারে আইকন সেট করা
            const container = document.getElementById("svgContainer");
            if (container) container.innerHTML = data.svg;
            if (resultArea) resultArea.style.display = "block";
            
            // ৩. আইকন সরাসরি স্লাইডে রিফ্লেক্ট করা
            insertToSlide(); 
            
        } else {
            alert("Error: " + (data.error || "Generation failed."));
        }
    } catch (error) {
        console.error("API Error:", error);
        alert("Server error. Please check your internet.");
    } finally {
        // লোডিং বন্ধ করা
        if (loadingDiv) loadingDiv.style.display = "none";
    }
}

async function insertToSlide() {
    const container = document.getElementById("svgContainer");
    const svgContent = container ? container.innerHTML : "";
    if (!svgContent) return;

    Office.context.document.setSelectedDataAsync(
        svgContent, 
        { coercionType: Office.CoercionType.XmlSvg },
        (result) => {
            if (result.status === Office.AsyncResultStatus.Failed) {
                console.error(result.error.message);
            }
        }
    );
}
