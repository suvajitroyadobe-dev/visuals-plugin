/* global Office */

Office.onReady((info) => {
    if (info.host === Office.HostType.PowerPoint) {
        document.getElementById("btnGenerateIcon").onclick = handleTextToIcon;
        document.getElementById("btnInsert").onclick = insertToSlide;
    }
});

async function handleTextToIcon() {
    const iconInput = document.getElementById("iconInput");
    const apiType = document.getElementById("apiType");
    const userApiKey = document.getElementById("userApiKey");

    const rawInput = iconInput ? iconInput.value : "";
    const selectedApi = apiType ? apiType.value : "huggingface";
    const apiKey = userApiKey ? userApiKey.value : "";
    const styleElement = document.querySelector('input[name="iconStyle"]:checked');
    const style = styleElement ? styleElement.value : "napkin";

    if (!rawInput) {
        alert("Please describe your icon first.");
        return;
    }

    // ১. ব্লিংকিং শুরু করা
    toggleLoading(true);
    document.getElementById("resultArea").style.display = "none";

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
            document.getElementById("svgContainer").innerHTML = data.svg;
            document.getElementById("resultArea").style.display = "block";
        } else {
            alert("Error: " + (data.error || "Generation failed."));
        }
    } catch (error) {
        console.error("API Error:", error);
        alert("Server error. Check if your Hugging Face Space is Running.");
    } finally {
        // ২. ব্লিংকিং বন্ধ করা
        toggleLoading(false);
    }
}

async function insertToSlide() {
    const svgContent = document.getElementById("svgContainer").innerHTML;
    if (!svgContent) return;
    Office.context.document.setSelectedDataAsync(svgContent, { coercionType: Office.CoercionType.XmlSvg });
}

function toggleLoading(isLoading) {
    const loadingDiv = document.getElementById("loading");
    if (loadingDiv) {
        loadingDiv.style.display = isLoading ? "block" : "none";
    }
}
