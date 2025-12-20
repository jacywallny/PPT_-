Office.onReady((info) => {
    if (info.host === Office.HostType.PowerPoint) {
        // 绑定按钮点击事件
        const btn = document.getElementById("invert-btn");
        if(btn) btn.onclick = invertSelectedImage;
    }
});

// 主入口函数
async function invertSelectedImage() {
    updateStatus("🔍 正在读取选中图片...", "blue");

    // 1. 获取选中的图片 (最大支持 4MB，防止内存溢出)
    Office.context.document.getSelectedDataAsync(
        Office.CoercionType.Image,
        { 
            imageLeft: 0, imageTop: 0, imageWidth: 0, imageHeight: 0,
            sliceSize: 4194304 // 4MB 切片，提高大图读取稳定性
        }, 
        function (result) {
            if (result.status === Office.AsyncResultStatus.Failed) {
                updateStatus("❌ 读取失败: 请确保你选中了一张图片！\n详细错误: " + result.error.message, "red");
                return;
            }
            
            // 获取到的数据是 Base64 字符串
            const imageBase64 = result.value;
            // 进入图片处理流程
            processImage(imageBase64);
        }
    );
}

// 图片处理核心逻辑
function processImage(base64Data) {
    updateStatus("⏳ 图片加载中...", "blue");

    const img = new Image();
    
    // 图片加载成功后的回调
    img.onload = function () {
        updateStatus("🎨 正在进行像素反色计算...", "blue");

        // ⚠️ 关键优化：使用 setTimeout 给 UI 一个喘息的机会，防止界面卡死
        setTimeout(() => {
            try {
                // 创建画布
                const canvas = document.createElement("canvas");
                const ctx = canvas.getContext("2d");
                
                // 这里的宽高决定了清晰度，保持原图大小
                canvas.width = img.width;
                canvas.height = img.height;

                // 将图片画到画布上
                ctx.drawImage(img, 0, 0);

                // 获取像素数据 (这是最耗时的步骤)
                const imageData = ctx.getImageData(0, 0, canvas.width, canvas.height);
                const data = imageData.data;

                // === 算法优化：遍历像素 ===
                // data[i] = R, data[i+1] = G, data[i+2] = B, data[i+3] = Alpha
                for (let i = 0; i < data.length; i += 4) {
                    data[i]     = 255 - data[i];     // Red
                    data[i + 1] = 255 - data[i + 1]; // Green
                    data[i + 2] = 255 - data[i + 2]; // Blue
                    // Alpha (透明度) 保持不变
                }

                // 将处理后的数据放回画布
                ctx.putImageData(imageData, 0, 0);

                // 导出为 Base64 (去除头部的 "data:image/png;base64,")
                const newBase64 = canvas.toDataURL("image/png").split(",")[1];
                
                // 替换 PPT 中的图片
                replaceImageInPPT(newBase64);

            } catch (error) {
                updateStatus("❌ 处理出错: " + error.message, "red");
                console.error(error);
            }
        }, 50); // 延时 50ms 执行，确保界面已刷新文字
    };

    // 图片加载失败的回调
    img.onerror = function() {
        updateStatus("❌ 图片数据解析失败，可能是图片格式不支持。", "red");
    };

    // 触发加载
    img.src = "data:image/png;base64," + base64Data;
}

// 将新图片回写到 PPT
function replaceImageInPPT(newBase64) {
    updateStatus("📤 正在替换原图...", "blue");

    Office.context.document.setSelectedDataAsync(
        newBase64,
        { coercionType: Office.CoercionType.Image },
        function (asyncResult) {
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                updateStatus("❌ 替换失败: " + asyncResult.error.message, "red");
            } else {
                updateStatus("✅ 成功！图片已反色。", "green");
            }
        }
    );
}

// 辅助函数：更新状态栏文字和颜色
function updateStatus(text, color) {
    const statusDiv = document.getElementById("status");
    if (statusDiv) {
        statusDiv.innerText = text;
        statusDiv.style.color = color || "black";
    }
}
