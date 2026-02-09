/* global Office */

// ১. পেজ লোড হওয়ামাত্রই বাটন খুঁজে ইভেন্ট সেট করা (Office.onReady-র বাইরে)
window.onload = function() {
    console.log("Window Loaded");
    const btn = document.getElementById("btnGenerateIcon");
    if (btn) {
        btn.onclick = handleTextToIcon;
    } else {
        console.error("Button not found!");
    }
    
    // ম্যানুয়াল ইনসার্ট বাটন
    const insBtn = document.getElementById("btnInsert");
    if (insBtn) insBtn.onclick = manualInsert;
};

Office.onReady((info) => {
    console.log("Office Ready: " + info.host);
});

async function handleTextToIcon() {
    // ২. বাটনে ক্লিক হয়েছে কি না তা চেক করা
    alert("Button Clicked! Starting Process..."); 

    const rawInput = document.getElementById("iconInput").value;
    const apiType = document.getElementById("apiType").value;
    const userApiKey = document.getElementById("userApiKey").value;
    let style = "napkin";
    const styleElem = document.querySelector('input[name="iconStyle"]:checked');
    if (styleElem) style = styleElem.value;

    if (!rawInput) {
        alert("Error: Please write a prompt first.");
        return;
    }

    // লোডিং দেখানো
    document.getElementById("loading").style.display = "block";
    
    try {
        // ৩. সার্ভারে রিকোয়েস্ট পাঠানো
        const response = await fetch("https://suvajit01-sr-visuals-backend.hf.space/generate-icon", {
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
        
        // ৪. সার্ভার থেকে ডাটা এসেছে কি না চেক করা
        if (data.raw_image) {
            alert("Success! Image received from server."); // ডিবাগ মেসেজ
            
            // প্রিভিউ সেট করা
            const container = document.getElementById("svgContainer");
            container.innerHTML = data.svg;
            container.dataset.rawImage = data.raw_image;
            document.getElementById("resultArea").style.display = "block";

            // ৫. স্লাইডে ছবি বসানো
            insertToSlide(data.raw_image);
        } else {
            alert("Server Error: " + JSON.stringify(data));
        }
    } catch (error) {
        alert("Network Error: " + error.message);
    } finally {
        document.getElementById("loading").style.display = "none";
    }
}

function insertToSlide(base64Code) {
    if (!base64Code) {
        alert("No image data to insert!");
        return;
    }

    Office.context.document.setSelectedDataAsync(
        base64Code,
        { coercionType: Office.CoercionType.Image },
        (result) => {
            if (result.status === Office.AsyncResultStatus.Failed) {
                alert("Insert Failed: " + result.error.message);
            } else {
                console.log("Image inserted!");
            }
        }
    );
}

function manualInsert() {
    const imgData = document.getElementById("svgContainer").dataset.rawImage;
    if (imgData) insertToSlide(imgData);
    else alert("Please generate an icon first.");
}
