/* global Office */

Office.onReady((info) => {
    if (info.host === Office.HostType.PowerPoint) {
        document.getElementById("btnGenerateIcon").onclick = handleTextToIcon;
        document.getElementById("btnInsert").onclick = insertToSlide;
    }
});

async function handleTextToIcon() {
    const rawInput = document.getElementById("iconInput").value;
    const apiType = document.getElementById("apiType").value;
    const userApiKey = document.getElementById("userApiKey").value;
    const styleElement = document.querySelector('input[name="iconStyle"]:checked');
    const style = styleElement ? styleElement.value : "napkin";

    if (!rawInput) {
        alert("Please describe your icon first.");
        return;
    }

    // লোডিং দেখানো এবং ব্লিংকিং শুরু
    document.getElementById("loading").style.display = "block";
    document.getElementById("resultArea").style.display = "none";

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
            document.getElementById("svgContainer").innerHTML = data.svg;
            document.getElementById("resultArea").style.display = "block";
        } else {
            alert("Error: " + (data.error || "Generation failed. Check your API Key."));
        }
    } catch (error) {
        console.error("API Error:", error);
        alert("Server error. Please check your internet or Space logs.");
    } finally {
        // লোডিং বন্ধ
        document.getElementById("loading").style.display = "none";
    }
}

async function insertToSlide() {
    const svgContent = document.getElementById("svgContainer").innerHTML;
    if (!svgContent) return;
    Office.context.document.setSelectedDataAsync(svgContent, { coercionType: Office.CoercionType.XmlSvg });
}
