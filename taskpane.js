/* global Office */

Office.onReady((info) => {
    if (info.host === Office.HostType.PowerPoint) {
        document.getElementById("btnGenerateIcon").onclick = handleTextToIcon;
        document.getElementById("btnInsert").onclick = insertToSlide;
    }
});

async function handleTextToIcon() {
    const rawInput = document.getElementById("iconInput")?.value || "";
    const apiType = document.getElementById("apiType")?.value || "huggingface";
    const userApiKey = document.getElementById("userApiKey")?.value || "";
    const styleElem = document.querySelector('input[name="iconStyle"]:checked');
    const style = styleElem ? styleElem.value : "napkin";

    if (!rawInput) {
        alert("Please describe your icon first.");
        return;
    }

    // লোডিং শুরু
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
            // প্রিভিউ দেখানো
            const container = document.getElementById("svgContainer");
            container.innerHTML = data.svg;
            document.getElementById("resultArea").style.display = "block";
            
            // অটোমেটিক ইনসার্ট
            insertToSlide(); 
        } else {
            alert("Error: " + (data.error || "Generation failed."));
        }
    } catch (error) {
        console.error("API Error:", error);
        alert("Server connection failed.");
    } finally {
        document.getElementById("loading").style.display = "none";
    }
}

function insertToSlide() {
    const container = document.getElementById("svgContainer");
    const svgContent = container ? container.innerHTML : "";
    
    if (!svgContent) return;

    // ১. চেষ্টা: সরাসরি ইমেজ হিসেবে ইনসার্ট করা (সবচেয়ে নিরাপদ পদ্ধতি)
    // আমরা SVG স্ট্রিং থেকে Base64 কোডটি বের করে আনব
    const base64Match = svgContent.match(/base64,([^"]*)/);
    
    if (base64Match && base64Match[1]) {
        const imageBase64 = base64Match[1]; // শুধু কোডটুকু
        
        Office.context.document.setSelectedDataAsync(
            imageBase64,
            { coercionType: Office.CoercionType.Image }, // Image টাইপ ব্যবহার করলে মিস হবে না
            (result) => {
                if (result.status === Office.AsyncResultStatus.Failed) {
                    console.error("Image Insert failed: " + result.error.message);
                    // ২. ব্যর্থ হলে HTML হিসেবে চেষ্টা
                    insertAsHtml(svgContent);
                }
            }
        );
    } else {
        // যদি Base64 না পাওয়া যায়, তবে সরাসরি HTML হিসেবে দিন
        insertAsHtml(svgContent);
    }
}

function insertAsHtml(htmlContent) {
    Office.context.document.setSelectedDataAsync(
        htmlContent,
        { coercionType: Office.CoercionType.Html },
        (result) => {
            if (result.status === Office.AsyncResultStatus.Failed) {
                console.error("HTML Insert failed: " + result.error.message);
                alert("Could not insert icon. Please try again.");
            }
        }
    );
}
এখানে ফাইল বিষয়বস্তু লিখুন
