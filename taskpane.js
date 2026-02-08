/* global Office */

Office.onReady((info) => {
    if (info.host === Office.HostType.PowerPoint) {
        document.getElementById("btnGenerateIcon").onclick = handleTextToIcon;
        document.getElementById("btnInsert").onclick = insertToSlide;
    }
});

async function handleTextToIcon() {
    // ১. সব ইনপুট এবং এলিমেন্ট সংগ্রহ করা
    const iconInput = document.getElementById("iconInput");
    const apiTypeElem = document.getElementById("apiType");
    const userApiKeyElem = document.getElementById("userApiKey");
    const loadingDiv = document.getElementById("loading");
    const resultArea = document.getElementById("resultArea");

    // ২. ভ্যালুগুলো সংগ্রহ এবং ভেরিফাই করা
    const rawInput = iconInput ? iconInput.value : "";
    const apiType = apiTypeElem ? apiTypeElem.value : "huggingface";
    const userApiKey = userApiKeyElem ? userApiKeyElem.value : "";
    const styleElem = document.querySelector('input[name="iconStyle"]:checked');
    const style = styleElem ? styleElem.value : "napkin";

    if (!rawInput) {
        alert("Please describe your icon first.");
        return;
    }

    // ৩. ব্লিংকিং এবং লোডিং শুরু করা
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
                api_type: apiType,
                user_key: userApiKey 
            }),
        });

        const data = await response.json();

        if (data.svg) {
            // ৪. প্রিভিউ দেখানো এবং সরাসরি স্লাইডে ইনসার্ট করা
            const container = document.getElementById("svgContainer");
            if (container) container.innerHTML = data.svg;
            if (resultArea) resultArea.style.display = "block";
            
            // স্লাইডে অটোমেটিক রিফ্লেক্ট করার জন্য
            insertToSlide(); 
            
        } else {
            alert("Error: " + (data.error || "Generation failed."));
        }
    } catch (error) {
        console.error("API Error:", error);
        alert("Server connection failed. Check your internet.");
    } finally {
        // ৫. লোডিং বন্ধ করা
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
                console.error("Insert failed: " + result.error.message);
            }
        }
    );
}
