Office.onReady((info) => {
    if (info.host === Office.HostType.PowerPoint) {
        const btn = document.getElementById("invert-btn");
        if(btn) btn.onclick = invertSelectedImage;
    }
});

async function invertSelectedImage() {
    updateStatus("🔍 正在读取选中图片...", "blue");

    Office.context.document.getSelectedDataAsync(
        Office.CoercionType.Image,
        { 
            sliceSize: 100000 // 这里的切片不用太大
        }, 
        function (result) {
            if (result.status === Office.AsyncResultStatus.Failed) {
                updateStatus("❌ 读取失败: 请确保选中了图片！\n" + result.error.message, "red");
                return;
            }
            
            const imageBase64 = result.value;
            
            // 🔍 智能侦测：检查数据是否有效
            if (!imageBase64 || imageBase64.length < 100) {
                updateStatus("❌ 错误: 无法获取图片数据，可能是矢量图或OLE对象。\n建议：请使用截图(Win+Shift+S)后粘贴再试。", "red");
                return;
            }
            
            processImage(imageBase64);
        }
    );
}

function processImage(base64Data) {
    updateStatus("⏳ 正在解析图片数据...", "blue");

    const img = new Image();
    
    img.onload = function () {
        updateStatus("🎨 正在反色处理...", "blue");
        setTimeout(() => {
            try {
                const canvas = document.createElement("canvas");
                const ctx = canvas.getContext("2d");
                canvas.width = img.width;
                canvas.height = img.height;

                ctx.drawImage(img, 0, 0);
                const imageData = ctx.getImageData(0, 0, canvas.width, canvas.height);
                const data = imageData.data;

                for (let i = 0; i < data.length; i += 4) {
                    data[i]     = 255 - data[i];
                    data[i + 1] = 255 - data[i + 1];
                    data[i + 2] = 255 - data[i + 2];
                }

                ctx.putImageData(imageData, 0, 0);
                const newBase64 = canvas.toDataURL("image/png").split(",")[1];
                replaceImageInPPT(newBase64);

            } catch (error) {
                updateStatus("❌ 算法错误: " + error.message, "red");
            }
        }, 50);
    };

    // 🚩 详细的错误诊断
    img.onerror = function() {
        // 打印前30个字符，看看是不是真正的图片数据
        const head = base64Data.substring(0, 30);
        updateStatus("❌ 格式不支持！\n浏览器无法识别此数据。\n数据头: " + head + "...\n👉 请尝试：Win+Shift+S 截图后再粘贴。", "red");
    };

    // 尝试添加 PNG 头加载
    img.src = "data:image/png;base64," + base64Data;
}

function replaceImageInPPT(newBase64) {
    updateStatus("📤 正在替换原图...", "blue");
    Office.context.document.setSelectedDataAsync(
        newBase64,
        { coercionType: Office.CoercionType.Image },
        function (asyncResult) {
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                updateStatus("❌ 替换失败: " + asyncResult.error.message, "red");
            } else {
                updateStatus("✅ 成功！", "green");
            }
        }
    );
}

function updateStatus(text, color) {
    const statusDiv = document.getElementById("status");
    if (statusDiv) {
        statusDiv.innerText = text;
        statusDiv.style.color = color || "black";
    }
}
