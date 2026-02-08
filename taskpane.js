/* global Office */

Office.onReady((info) => {
    if (info.host === Office.HostType.PowerPoint) {
        document.getElementById("btnGenerateIcon").onclick = handleTextToIcon;
        document.getElementById("btnInsert").onclick = insertToSlide;
    }
});

// --- ১. টেক্সট টু আইকন (ইউজার API Key সহ) ---
// taskpane.js এর handleTextToIcon ফাংশনে পরিবর্তন
const apiType = document.getElementById("apiType").value;
const userApiKey = document.getElementById("userApiKey").value;

const response = await fetch(serverUrl, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ 
        prompt: rawInput,
        style: style,
        api_type: apiType, // কোন এপিআই ব্যবহার হবে
        user_key: userApiKey 
    }),
});

    // লোডিং দেখানো
    toggleLoading(true);
    document.getElementById("resultArea").style.display = "none";

    try {
        const serverUrl = "https://suvajit01-sr-visuals-backend.hf.space/generate-icon"; // আপনার ব্যাকএন্ড এন্ডপয়েন্ট
        const style = document.querySelector('input[name="iconStyle"]:checked').value;

        const response = await fetch(serverUrl, {
            method: "POST",
            headers: { 
                "Content-Type": "application/json"
            },
            body: JSON.stringify({ 
                prompt: rawInput,
                style: style,
                user_key: userApiKey // ইউজারের কি ব্যাকএন্ডে পাঠানো হচ্ছে
            }),
        });

        const data = await response.json();

        if (data.svg) {
            const container = document.getElementById("svgContainer");
            container.innerHTML = data.svg;
            document.getElementById("resultArea").style.display = "block";
        } else {
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
        await PowerPoint.run(async (context) => {
            const sheet = context.presentation.slides.getItemAt(0); // প্রথম স্লাইডে ইনসার্ট
            // স্লাইডে ইমেজ হিসেবে ইনসার্ট করার জন্য Office.js এর সাহায্য নেওয়া
            Office.context.document.setSelectedDataAsync(
                svgContent,
                { coercionType: Office.CoercionType.XmlSvg },
                (result) => {
                    if (result.status === Office.AsyncResultStatus.Failed) {
                        console.error(result.error.message);
                    }
                }
            );
        });
    } catch (error) {
        // যদি সরাসরি SVG সাপোর্ট না করে, তবে বেসিক মেথড
        Office.context.document.setSelectedDataAsync(svgContent, { coercionType: Office.CoercionType.XmlSvg });
    }
}

// --- ৩. লোডিং এবং স্ট্যাটাস কন্ট্রোল ---
function toggleLoading(isLoading) {
    const loadingDiv = document.getElementById("loading");
    loadingDiv.style.display = isLoading ? "block" : "none";
}

// স্ট্যাটাস রিপোর্ট ফাংশন (পুরানো কোডের সাথে সামঞ্জস্য রাখতে)
function reportStatus(message, isBusy) {
    console.log(message);
}

