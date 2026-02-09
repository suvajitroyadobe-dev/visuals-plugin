/* global Office */

Office.onReady((info) => {
    if (info.host === Office.HostType.PowerPoint) {
        // বাটন সেটআপ
        const genBtn = document.getElementById("btnGenerateIcon");
        if (genBtn) genBtn.onclick = handleTextToIcon;
        
        const insertBtn = document.getElementById("btnInsert");
        if (insertBtn) insertBtn.onclick = manualInsert;

        const vecBtn = document.getElementById("btnVectorize");
        if (vecBtn) vecBtn.onclick = handleImageToVector;
    }
});

// ১. আইকন জেনারেট ফাংশন
async function handleTextToIcon() {
    const rawInput = document.getElementById("iconInput").value;
    const apiType = document.getElementById("apiType").value;
    const userApiKey = document.getElementById("userApiKey").value;
    
    // রেডিও বাটন সিলেকশন চেক
    let style = "napkin";
    const styleElem = document.querySelector('input[name="iconStyle"]:checked');
    if (styleElem) style = styleElem.value;

    if (!rawInput) {
        alert("Please describe your icon first.");
        return;
    }

    // লোডিং শুরু
    document.getElementById("loading").style.display = "block";
    document.getElementById("resultArea").style.display = "none";

    try {
        console.log("Sending Request...");
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
        console.log("Response:", data);

        if (data.svg) {
            // প্রিভিউ সেট করা
            document.getElementById("svgContainer").innerHTML = data.svg;
            document.getElementById("resultArea").style.display = "block";
            
            // অটোমেটিক ইনসার্ট চেষ্টা করা
            insertToSlide(data.svg);
        } else {
            alert("Error from Server: " + (data.error || "Unknown Error"));
        }
    } catch (error) {
        console.error(error);
        alert("Connection Failed! Check console for details.");
    } finally {
        document.getElementById("loading").style.display = "none";
    }
}

// ২. ইমেজ টু ভেক্টর ফাংশন
async function handleImageToVector() {
    const fileInput = document.getElementById("imageInput");
    if (!fileInput.files.length) {
        alert("Please select an image.");
        return;
    }

    document.getElementById("loading").style.display = "block";
    const formData = new FormData();
    formData.append("file", fileInput.files[0]);

    try {
        const response = await fetch("https://suvajit01-sr-visuals-backend.hf.space/image-to-vector", {
            method: "POST",
            body: formData
        });
        const data = await response.json();
        
        if (data.svg) {
            document.getElementById("svgContainer").innerHTML = data.svg;
            document.getElementById("resultArea").style.display = "block";
            insertToSlide(data.svg);
        }
    } catch (e) {
        alert("Vector Error: " + e.message);
    } finally {
        document.getElementById("loading").style.display = "none";
    }
}

// ৩. ম্যানুয়াল ইনসার্ট বাটন ফাংশন
function manualInsert() {
    const content = document.getElementById("svgContainer").innerHTML;
    if(content) insertToSlide(content);
    else alert("No icon to insert!");
}

// ৪. মেইন ইনসার্ট লজিক (সবচেয়ে গুরুত্বপূর্ণ অংশ)
function insertToSlide(svgString) {
    if (!svgString) return;

    // ক) প্রথমে দেখব এটা ইমেজ-বেসড SVG কি না (Text-to-Icon এর জন্য)
    const base64Match = svgString.match(/base64,([^"']+)/);
    
    if (base64Match && base64Match[1]) {
        console.log("Trying to insert as Image...");
        
        Office.context.document.setSelectedDataAsync(
            base64Match[1], // শুধু বেস৬৪ কোড
            { coercionType: Office.CoercionType.Image },
            (result) => {
                if (result.status === Office.AsyncResultStatus.Failed) {
                    console.error("Image Insert Failed:", result.error.message);
                    // ইমেজ ফেইল করলে প্ল্যান বি: HTML
                    tryHtmlInsert(svgString);
                } else {
                    console.log("Success: Image Inserted!");
                }
            }
        );
    } else {
        // খ) যদি ইমেজ না থাকে (পিওর ভেক্টর), তবে SVG বা HTML হিসেবে চেষ্টা করব
        console.log("Trying to insert as Vector/HTML...");
        tryHtmlInsert(svgString);
    }
}

function tryHtmlInsert(content) {
    Office.context.document.setSelectedDataAsync(
        content,
        { coercionType: Office.CoercionType.Html },
        (result) => {
            if (result.status === Office.AsyncResultStatus.Failed) {
                // এখানে আমরা আসল কারণ দেখতে পাব
                alert("Final Insert Failed. Error: " + result.error.message);
            }
        }
    );
}
