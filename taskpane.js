/* global Office */

Office.onReady((info) => {
    if (info.host === Office.HostType.PowerPoint) {
        // বাটন ক্লিক ইভেন্ট সেট করা
        const generateBtn = document.getElementById("btnGenerateIcon");
        if (generateBtn) {
            generateBtn.onclick = handleTextToIcon;
        }
        
        const insertBtn = document.getElementById("btnInsert");
        if (insertBtn) {
            insertBtn.onclick = insertToSlide;
        }
    }
});

async function handleTextToIcon() {
    // ১. এলিমেন্টগুলো থেকে সরাসরি ভ্যালু নেওয়া
    const rawInput = document.getElementById("iconInput")?.value || "";
    const apiType = document.getElementById("apiType")?.value || "huggingface";
    const userApiKey = document.getElementById("userApiKey")?.value || "";
    const styleElem = document.querySelector('input[name="iconStyle"]:checked');
    const style = styleElem ? styleElem.value : "napkin";

    if (!rawInput) {
        alert("Please describe your icon first.");
        return;
    }

    // ২. লোডিং এবং ব্লিংকিং শুরু করা
    const loadingDiv = document.getElementById("loading");
    const resultArea = document.getElementById("resultArea");
    
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
            // ৩. আইকন প্রিভিউ দেখানো
            const container = document.getElementById("svgContainer");
            if (container) container.innerHTML = data.svg;
            if (resultArea) resultArea.style.display = "block";
            
            // ৪. স্লাইডে অটোমেটিক পাঠানো
            insertToSlide(); 
            
        } else {
            alert("Error: " + (data.error || "Generation failed."));
        }
    } catch (error) {
        console.error("API Error:", error);
        alert("Server connection failed.");
    } finally {
        // লোডিং বন্ধ করা
        if (loadingDiv) loadingDiv.style.display = "none";
    }
}

async function insertToSlide() {
    const container = document.getElementById("svgContainer");
    // innerHTML এর বদলে firstElementChild ব্যবহার করা নিরাপদ
    const svgElement = container ? container.firstElementChild : null;
    
    if (!svgElement) return;

    // SVG এলিমেন্টকে স্ট্রিং এ রূপান্তর করা
    const serializer = new XMLSerializer();
    const svgContent = serializer.serializeToString(svgElement);

    Office.context.document.setSelectedDataAsync(
        svgContent, 
        { coercionType: Office.CoercionType.XmlSvg },
        (result) => {
            if (result.status === Office.AsyncResultStatus.Failed) {
                console.error("Insert failed: " + result.error.message);
                // যদি XmlSvg ফেল করে তবে ব্যাকআপ হিসেবে HTML ট্রাই করুন
                Office.context.document.setSelectedDataAsync(svgContent, { coercionType: Office.CoercionType.Html });
            }
        }
    );
}

