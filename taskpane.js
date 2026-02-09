/* global Office */

Office.onReady((info) => {
    if (info.host === Office.HostType.PowerPoint) {
        // বাটন ইভেন্ট কানেক্ট করা
        document.getElementById("btnGenerateIcon").onclick = handleTextToIcon;
        
        // ইমেজ টু ভেক্টর বাটন (যদি ভবিষ্যতে ব্যবহার করেন)
        const vecBtn = document.getElementById("btnVectorize");
        if (vecBtn) vecBtn.onclick = handleImageToVector;

        // ম্যানুয়াল ইনসার্ট বাটন
        document.getElementById("btnInsert").onclick = () => {
            const svgContent = document.getElementById("svgContainer").innerHTML;
            if (svgContent) insertToSlide(svgContent);
        };
    }
});

async function handleTextToIcon() {
    const rawInput = document.getElementById("iconInput").value;
    const apiType = document.getElementById("apiType").value;
    const userApiKey = document.getElementById("userApiKey").value;
    // রেডিও বাটন চেক
    const styleElem = document.querySelector('input[name="iconStyle"]:checked');
    const style = styleElem ? styleElem.value : "napkin";

    if (!rawInput) {
        document.getElementById("iconInput").style.border = "2px solid red";
        return;
    } else {
        document.getElementById("iconInput").style.border = "1px solid #ccc";
    }

    // লোডিং স্ক্রিন দেখানো
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
            // প্রিভিউ সেট করা
            const container = document.getElementById("svgContainer");
            container.innerHTML = data.svg;
            
            document.getElementById("resultArea").style.display = "block";
            
            // অটোমেটিক স্লাইডে পাঠানো
            insertToSlide(data.svg);
        } else {
            console.error("API Error:", data);
            alert("Error: " + (data.error || "Failed to generate"));
        }
    } catch (error) {
        console.error("Network Error:", error);
    } finally {
        document.getElementById("loading").style.display = "none";
    }
}

// ইমেজ টু ভেক্টর হ্যান্ডলার (Img2Vec ট্যাবের জন্য)
async function handleImageToVector() {
    const fileInput = document.getElementById("imageInput");
    if (!fileInput.files.length) {
        alert("Please select an image first.");
        return;
    }

    document.getElementById("loading").style.display = "block";
    document.getElementById("resultArea").style.display = "none";

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
        } else {
            alert("Vector Error: " + data.error);
        }
    } catch (e) {
        console.error(e);
    } finally {
        document.getElementById("loading").style.display = "none";
    }
}

function insertToSlide(svgString) {
    if (!svgString) return;

    const base64Match = svgString.match(/base64,([^"']+)/);
    
    if (base64Match && base64Match[1]) {
        // ইমেজ হিসেবে ইনসার্ট
        Office.context.document.setSelectedDataAsync(
            base64Match[1],
            { coercionType: Office.CoercionType.Image },
            (result) => {
                if (result.status === Office.AsyncResultStatus.Failed) {
                    insertAsHtml(svgString);
                }
            }
        );
    } else {
        // ভেক্টর বা HTML হিসেবে ইনসার্ট
        insertAsHtml(svgString);
    }
}

function insertAsHtml(htmlContent) {
    Office.context.document.setSelectedDataAsync(
        htmlContent,
        { coercionType: Office.CoercionType.Html },
        (res) => { if (res.status === 'failed') console.error(res.error.message); }
    );
}
