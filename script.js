Office.onReady((info) => {
    if (info.host === Office.HostType.PowerPoint) {
        document.getElementById("invert-btn").onclick = invertSelectedImage;
    }
});

async function invertSelectedImage() {
    const statusDiv = document.getElementById("status");
    statusDiv.innerText = "正在读取图片...";

    // 1. 获取选中的图片数据 (Base64格式)
    Office.context.document.getSelectedDataAsync(
        Office.CoercionType.Image,
        { imageLeft: 0, imageTop: 0, imageWidth: 0, imageHeight: 0 }, // 保持原比例
        function (result) {
            if (result.status === Office.AsyncResultStatus.Failed) {
                statusDiv.innerText = "错误：请确保选中了一张图片！";
                return;
            }

            const imageBase64 = result.value;
            processImage(imageBase64);
        }
    );
}

function processImage(base64Data) {
    const statusDiv = document.getElementById("status");
    statusDiv.innerText = "正在处理像素...";

    const img = new Image();
    img.onload = function () {
        // 创建画布
        const canvas = document.createElement("canvas");
        const ctx = canvas.getContext("2d");
        canvas.width = img.width;
        canvas.height = img.height;

        // 画入原图
        ctx.drawImage(img, 0, 0);

        // 获取像素数据
        const imageData = ctx.getImageData(0, 0, canvas.width, canvas.height);
        const data = imageData.data;

        // 遍历所有像素进行反色 (Invert)
        for (let i = 0; i < data.length; i += 4) {
            data[i] = 255 - data[i];     // Red
            data[i + 1] = 255 - data[i + 1]; // Green
            data[i + 2] = 255 - data[i + 2]; // Blue
            // data[i+3] 是 Alpha (透明度)，保持不变
        }

        // 放回画布
        ctx.putImageData(imageData, 0, 0);

        // 导出新图片并替换 PPT 中的选中项
        const newBase64 = canvas.toDataURL("image/png").split(",")[1];
        replaceImageInPPT(newBase64);
    };
    
    // 加载图片
    img.src = "data:image/png;base64," + base64Data;
}

function replaceImageInPPT(newBase64) {
    const statusDiv = document.getElementById("status");
    statusDiv.innerText = "正在替换图片...";

    Office.context.document.setSelectedDataAsync(
        newBase64,
        { coercionType: Office.CoercionType.Image },
        function (asyncResult) {
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                statusDiv.innerText = "替换失败: " + asyncResult.error.message;
            } else {
                statusDiv.innerText = "✅ 反色完成！";
            }
        }
    );
}
