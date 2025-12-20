Office.onReady((info) => {
    if (info.host === Office.HostType.PowerPoint) {
        // 绑定新按钮的点击事件
        const btn = document.getElementById("paste-btn");
        if(btn) btn.onclick = handlePasteAndInvert;
    }
});

// === 主入口：处理粘贴并反色 ===
async function handlePasteAndInvert() {
    updateStatus("🔍 正在读取剪贴板...", "blue");

    try {
        // 1. 请求读取剪贴板内容
        // 注意：第一次运行时，浏览器可能会在顶部弹窗询问“是否允许访问剪贴板”，请点击允许。
        const clipboardItems = await navigator.clipboard.read();

        let imageBlob = null;

        // 2. 遍历剪贴板项目，寻找图片格式
        for (const item of clipboardItems) {
            // 优先寻找 png，其次 jpeg
            if (item.types.includes("image/png")) {
                imageBlob = await item.getType("image/png");
                break;
            } else if (item.types.includes("image/jpeg")) {
                imageBlob = await item.getType("image/jpeg");
                break;
            }
        }

        if (!imageBlob) {
            updateStatus("❌ 剪贴板里没有发现图片！\n请先在 PPT 中选中对象并按下 Ctrl+C。", "red");
            return;
        }

        // 3. 将图片二进制 Blob 转换为 Base64 供后续处理
        updateStatus("⏳ 获取到图片，准备处理...", "blue");
        const base64Data = await blobToBase64(imageBlob);
        
        // 4. 进入核心反色流程 (复用之前的逻辑)
        processImage(base64Data);

    } catch (err) {
        // 捕获权限错误或其他异常
        console.error(err);
        if (err.name === 'NotAllowedError') {
             updateStatus("❌ 无法读取剪贴板。\n请确保您在浏览器提示时点击了“允许”访问剪贴板。", "red");
        } else {
             updateStatus("❌ 读取剪贴板出错:\n" + err.message, "red");
        }
    }
}


// === 核心：图片反色逻辑 (复用之前稳定版的代码) ===
function processImage(base64DataNoPrefix) {
    updateStatus("🎨 正在进行像素反色计算...", "blue");
    
    // 需要加上前缀才能让 Image 对象识别
    const fullBase64Str = "data:image/png;base64," + base64DataNoPrefix;

    const img = new Image();
    
    img.onload = function () {
        // 使用 setTimeout 防止界面卡死
        setTimeout(() => {
            try {
                const canvas = document.createElement("canvas");
                const ctx = canvas.getContext("2d");
                canvas.width = img.width;
                canvas.height = img.height;

                ctx.drawImage(img, 0, 0);
                const imageData = ctx.getImageData(0, 0, canvas.width, canvas.height);
                const data = imageData.data;

                // 像素反色循环
                for (let i = 0; i < data.length; i += 4) {
                    data[i]     = 255 - data[i];     // R
                    data[i + 1] = 255 - data[i + 1]; // G
                    data[i + 2] = 255 - data[i + 2]; // B
                }

                ctx.putImageData(imageData, 0, 0);
                // 导出新图片 Base64 (去除前缀用于 PPT 插入)
                const newBase64 = canvas.toDataURL("image/png").split(",")[1];
                
                insertImageIntoPPT(newBase64);

            } catch (error) {
                updateStatus("❌ 处理出错: " + error.message, "red");
            }
        }, 50);
    };

    img.onerror = function() {
        updateStatus("❌ 剪贴板中的数据不是有效的图片格式。", "red");
    };

    img.src = fullBase64Str;
}


// === 将新图片插入 PPT ===
function insertImageIntoPPT(newBase64) {
    updateStatus("📤 正在插入反色后的图片...", "blue");
    
    // 使用 setSelectedDataAsync 插入图片
    // 如果当前有选中内容，会被替换；如果没有，则插入到光标位置。
    Office.context.document.setSelectedDataAsync(
        newBase64,
        { coercionType: Office.CoercionType.Image },
        function (asyncResult) {
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                updateStatus("❌ 插入失败: " + asyncResult.error.message, "red");
            } else {
                updateStatus("✅ 成功！反色图片已插入。", "green");
            }
        }
    );
}


// === 辅助工具：Blob 转 Base64 ===
function blobToBase64(blob) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onloadend = () => {
            // result 是类似 "data:image/png;base64,XXXX" 的字符串
            // 我们只需要逗号后面的部分
            const base64Raw = reader.result.split(',')[1];
            resolve(base64Raw);
        };
        reader.onerror = reject;
        reader.readAsDataURL(blob);
    });
}

// === 辅助工具：更新状态栏 ===
function updateStatus(text, color) {
    const statusDiv = document.getElementById("status");
    if (statusDiv) {
        statusDiv.innerText = text;
        statusDiv.style.color = color || "black";
    }
}
