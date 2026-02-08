/* global Office */

Office.onReady((info) => {
    if (info.host === Office.HostType.PowerPoint) {
        document.getElementById("btnGenerateIcon").onclick = handleTextToIcon;
        document.getElementById("btnInsert").onclick = insertToSlide;
    }
});

// --- ১. টেক্সট টু আইকন লজিক ---
async function handleTextToIcon() {
    const rawInput = document.getElementById("iconInput").value;
    const apiType = document.getElementById("apiType").value;
    const userApiKey = document.getElementById("userApiKey").value;
    const style = document.querySelector('input[name="iconStyle"]:checked').value;

    if (!rawInput) {
        alert("Please describe your icon first.");
        return;
    }

    // লোডিং স্ক্রিন দেখানো
    toggleLoading(true);
    document.getElementById("resultArea").style.display = "none";

    try {
        const serverUrl = "https://suvajit01-sr-visuals-backend.hf.space/generate-icon";
        
        const response = await fetch(serverUrl, {
            method: "POST",
            headers: { 
                "Content-Type": "application/json" 
            },
            body: JSON.stringify({ 
                prompt: rawInput,
                style: style,
                api_type: apiType, // ইউজার কি এপিআই টাইপ সিলেক্ট করেছে
                user_key: userApiKey // ইউজারের কি ব্যাকএন্ডে পাঠানো হচ্ছে
            }),
        });

        const data = await response.json();

        if (data.svg) {
            const container = document.getElementById("svgContainer");
            container.innerHTML = data.svg;
            document.getElementById("resultArea").style.display = "block";
        } else {
            // এরর মেসেজ হ্যান্ডলিং
            alert("Error: " + (data.error || "Could not generate icon. Check your API Key."));
        }
    } catch (error) {
        console.error("API Error:", error);
        alert("Server error. Please make sure your Hugging Face Space is Running.");
    } finally {
        toggleLoading(false);
    }
}

// --- ২. স্লাইডে ইনসার্ট লজিক ---
async function insertToSlide() {
    const svgContent = document.getElementById("svgContainer").innerHTML;
    if (!svgContent) return;

    try {
        Office.context.document.setSelectedDataAsync(
            svgContent,
            { coercionType: Office.CoercionType.XmlSvg },
            (result) => {
                if (result.status === Office.AsyncResultStatus.Failed) {
                    console.error(result.error.message);
                }
            }
        );
    } catch (error) {
        console.error("Insert Error:", error);
    }
}

// --- ৩. লোডিং কন্ট্রোল ---
function toggleLoading(isLoading) {
    const loadingDiv = document.getElementById("loading");
    if (loadingDiv) {
        loadingDiv.style.display = isLoading ? "block" : "none";
    }
}
