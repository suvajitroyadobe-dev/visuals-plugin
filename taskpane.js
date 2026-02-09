/* global Office */

Office.onReady((info) => {
    if (info.host === Office.HostType.PowerPoint) {
        // বাটন ইভেন্ট হ্যান্ডলার
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
    // ১. ইনপুট ভ্যালু সংগ্রহ
    const rawInput = document.getElementById("iconInput")?.value || "";
    const apiType = document.getElementById("apiType")?.value || "huggingface";
    const userApiKey = document.getElementById("userApiKey")?.value || "";
    const styleElem = document.querySelector('input[name="iconStyle"]:checked');
    const style = styleElem ? styleElem.value : "napkin";

    if (!rawInput) {
        alert("Please describe your icon first.");
        return;
    }

    // ২. লোডিং দেখানো
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
            // ৩. প্রিভিউ দেখানো
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
        alert("Server connection failed. Check your internet.");
    } finally {
        if (loadingDiv) loadingDiv.style.display = "none";
    }
}

async function insertToSlide() {
    const container = document.getElementById("svgContainer");
    const svgContent = container ? container.innerHTML : "";
    
    // সংশোধিত লাইন (বাংলা লেখাটি বাদ দেওয়া হয়েছে)
    if (!svgContent) return;

    // PowerPoint-এ Image যুক্ত SVG দেখানোর জন্য 'Html' টাইপ ব্যবহার করা হচ্ছে
    Office.context.document.setSelectedDataAsync(
        svgContent,
        { coercionType: Office.CoercionType.Html },
        (result) => {
            if (result.status === Office.AsyncResultStatus.Failed) {
                console.error("Insert failed: " + result.error.message);
                // যদি Html কাজ না করে, তবে ইউজারের ম্যানুয়ালি কপি করার অপশন থাকেই
                alert("Could not insert automatically. Please copy the icon manually.");
            }
        }
    );
}
এখানে ফাইল বিষয়বস্তু লিখুন
