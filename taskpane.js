/*
 * SR Visuals - Main Logic
 * Connects UI, ImageTracer, and Hugging Face Backend
 */

Office.onReady((info) => {
    if (info.host === Office.HostType.PowerPoint) {
        // ইভেন্ট লিসেনার যোগ করা
        document.getElementById("btnGenerateIcon").onclick = handleTextToIcon;
        document.getElementById("btnVectorize").onclick = handleImageToVector;
        document.getElementById("imageUpload").onchange = showImagePreview;
        document.getElementById("btnInsert").onclick = insertSvgToSlide;
    }
});

// --- ১. গ্লোবাল ভেরিয়েবল ---
let generatedSVG = ""; // তৈরি হওয়া SVG এখানে জমা থাকবে

// --- ২. সার্ভার কানেকশন (অনুবাদ) ---
async function translateWithMyServer(text) {
    // আপনার তৈরি করা সার্ভার
    const serverUrl = "https://suvajit01-sr-visuals-backend.hf.space/translate";
    
    try {
        const response = await fetch(serverUrl, {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({
                text: text,
                source_lang: "auto",
                target_lang: "en"
            }),
        });

        const data = await response.json();
        return data.translated || text; // অনুবাদ না হলে আসল টেক্সট
    } catch (error) {
        console.error("Server error:", error);
        return text; // সার্ভার অফ থাকলে আসল টেক্সট
    }
}

// --- ৩. ইমেজ টু ভেক্টর (আপনার নতুন ফিচার) ---
function showImagePreview() {
    const file = document.getElementById("imageUpload").files[0];
    const preview = document.getElementById("uploadedPreview");
    const placeholder = document.getElementById("uploadPlaceholder");

    if (file) {
        const reader = new FileReader();
        reader.onload = function(e) {
            preview.src = e.target.result;
            preview.style.display = "block";
            placeholder.style.display = "none";
        };
        reader.readAsDataURL(file);
    }
}

function handleImageToVector() {
    const file = document.getElementById("imageUpload").files[0];
    if (!file) {
        reportStatus("Please select an image first!");
        return;
    }

    reportStatus("Converting to Vector...", true);

    // ImageTracer ব্যবহার করে ছবি ট্রেস করা
    const reader = new FileReader();
    reader.onload = function(event) {
        const imgData = event.target.result;
        
        // ImageTracer কনফিগারেশন (সুন্দর লাইনের জন্য)
        const options = {
            ltres: 1,       // Line tolerance (lower = more detailed)
            qtres: 1,       // Curve tolerance
            scale: 1,       // Scale factor
            strokewidth: 2, // লাইন কতটা মোটা হবে
            linefilter: true // নয়েজ কমানো
        };

        // কনভার্ট করা হচ্ছে
        ImageTracer.imageToSVG(imgData, function(svgstr) {
            displaySVG(svgstr);
            reportStatus(""); // লোডিং বন্ধ
        }, options);
    };
    reader.readAsDataURL(file);
}

// --- ৪. টেক্সট টু আইকন (আসল AI জেনারেশন) ---
async function handleTextToIcon() {
    const rawInput = document.getElementById("iconInput").value;
    if (!rawInput) return;

    reportStatus("Processing...", true);

    // ১. আপনার সার্ভার দিয়ে অনুবাদ
    const translatedPrompt = await translateWithMyServer(rawInput);
    
    reportStatus("Generating your visual...", true);
    
    try {
        // ২. আপনার Hugging Face ব্যাকএন্ডে রিকোয়েস্ট পাঠানো
        const serverUrl = "https://suvajit01-sr-visuals-backend.hf.space/generate-icon"; 
        
        const response = await fetch(serverUrl, {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({ 
                prompt: translatedPrompt,
                style: document.querySelector('input[name="iconStyle"]:checked').value 
            }),
        });

        const data = await response.json();

        if (data.svg) {
            // আসল জেনারেটেড SVG দেখানো
            displaySVG(data.svg);
            reportStatus(`Success! Created: "${translatedPrompt}"`);
        } else {
            reportStatus("Error: API is busy or failing.");
        }
    } catch (error) {
        console.error("API Error:", error);
        reportStatus("Backend is starting up. Please try again in 30 seconds.");
    }
}

// --- ৫. হেল্পার ফাংশন ---
function displaySVG(svgString) {
    generatedSVG = svgString;
    const container = document.getElementById("svgContainer");
    container.innerHTML = svgString;
    
    // SVG এর সাইজ ঠিক করা যাতে প্রিভিউতে দেখা যায়
    const svgEl = container.querySelector("svg");
    if(svgEl) {
        svgEl.setAttribute("width", "100%");
        svgEl.setAttribute("height", "100%");
    }

    document.getElementById("resultArea").style.display = "block";
}

function insertSvgToSlide() {
    if (!generatedSVG) return;

    Office.context.document.setSelectedDataAsync(
        generatedSVG,
        { coercionType: Office.CoercionType.XmlSvg },
        (asyncResult) => {
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                console.error(asyncResult.error.message);
            }
        }
    );
}

function reportStatus(message, isLoading = false) {
    const loader = document.getElementById("loading");
    if (isLoading) {
        loader.style.display = "block";
        loader.querySelector("p").innerText = message;
    } else {
        loader.style.display = "none";
        // এরর বা সাধারণ মেসেজ থাকলে অ্যালার্ট দেখানো যেতে পারে
        if(message) console.log(message);
    }

}

