/* global Office, Word, PowerPoint */

Office.onReady(() => {
  const btn = document.getElementById("runBtn");
  if (btn) btn.onclick = runInvert;
});

async function runInvert() {
  updateStatus("⏳ 正在处理...");

  try {
    if (Office.context.host === Office.HostType.Word) {
      await runInvertInWord();
      return;
    }

    if (Office.context.host === Office.HostType.PowerPoint) {
      await runInvertInPowerPoint();
      return;
    }

    // 其它宿主走通用
    runInvertCommon();
  } catch (e) {
    console.error(e);
    updateStatus("❌ 发生异常：" + (e?.message || e));
  }
}

/* =========================
 * PowerPoint：稳定路径
 * ========================= */
async function runInvertInPowerPoint() {
  // 1) 能力检测：如果 PPT API 不支持，就不要硬跑（否则你会看到各种“枚举不支持”之类报错）
  const hasPptApi18 = Office.context.requirements.isSetSupported("PowerPointApi", "1.8");
  const hasPptApi110 = Office.context.requirements.isSetSupported("PowerPointApi", "1.10");

  // ImageCoercion 通常用于通用 getSelectedDataAsync(Image)
  const hasImageCoercion = Office.context.requirements.isSetSupported("ImageCoercion", "1.2");

  // 如果连 1.10 都没有，基本无法“确保”对选中图片做导出→反色→写回
  if (!hasPptApi110) {
    updateStatus(
      "❌ 当前 PowerPoint 环境不支持 PowerPointApi 1.10。\n" +
      "这意味着无法使用 getImageAsBase64 导出选中图片/形状，因此无法保证反色成功。\n\n" +
      "检测结果：\n" +
      `- PowerPointApi 1.8: ${hasPptApi18}\n` +
      `- PowerPointApi 1.10: ${hasPptApi110}\n` +
      `- ImageCoercion 1.2: ${hasImageCoercion}\n\n` +
      "建议：使用 Microsoft 365 桌面版 PowerPoint（Win/Mac）并更新到较新版本。\n" +
      "我也会尝试通用模式（成功率取决于环境）。"
    );

    // 尝试通用模式（可能仍失败）
    try { runInvertCommon(); } catch (_) {}
    return;
  }

  // 2) PPT 专用路径：选中 shape -> 导出 base64 -> 反色 -> 写回 fill
  try {
    await PowerPoint.run(async (context) => {
      const shapes = context.presentation.getSelectedShapes();
      const count = shapes.getCount();

      await context.sync();

      if (!count || count.value === 0) {
        updateStatus("❌ 未检测到选中的对象。\n请在幻灯片中单击选中图片本体（出现 8 个控制点）后再点击按钮。");
        return;
      }

      shapes.load("items");
      await context.sync();

      updateStatus(`🎨 已选中 ${count.value} 个对象，正在反色...`);

      // 逐个处理
      for (const shape of shapes.items) {
        // 导出渲染图（PNG base64，不带 data: 前缀）
        const imgResult = shape.getImageAsBase64({ format: "Png" });
        await context.sync();

        const base64 = imgResult.value;
        if (!base64) continue;

        // 反色（输出为 data:image/png;base64,xxxx）
        const newBase64DataUrl = await invertImagePromise(base64);

        // setImage 需要纯 base64（不含 data:image/... 前缀）
        const cleanBase64 = newBase64DataUrl.split(",")[1];

        // 写回：将形状填充设置为图片
        shape.fill.setImage(cleanBase64);
      }

      await context.sync();
      updateStatus("✅ 成功！已反色");
    });
  } catch (e) {
    console.error(e);

    // 把最关键的信息吐给你（你截图里那个“枚举不支持”就是这里来的）
    const msg = e?.message || String(e);

    updateStatus(
      "❌ PowerPoint 专用模式失败。\n" +
      "错误信息：\n" + msg + "\n\n" +
      "说明：若出现“当前宿主应用程序中不支持枚举/不支持此 API”等提示，通常是 PowerPoint 环境不支持所需 API。\n" +
      "我将尝试通用模式（成功率取决于环境）。"
    );

    try { runInvertCommon(); } catch (_) {}
  }
}

/* =========================
 * Word：嵌入式图片强力路径
 * ========================= */
async function runInvertInWord() {
  try {
    await Word.run(async (context) => {
      const selection = context.document.getSelection();
      const pictures = selection.inlinePictures;

      pictures.load("items");
      await context.sync();

      if (pictures.items.length === 0) {
        updateStatus("❌ 未检测到嵌入型图片。\n请右键图片 → 文字环绕 → 设为【嵌入型】后重试。");
        return;
      }

      const pic = pictures.items[0];

      // 读取 base64
      const base64Result = pic.getBase64ImageSrc();
      await context.sync();

      const base64 = base64Result.value;
      if (!base64) {
        updateStatus("❌ 无法读取图片数据");
        return;
      }

      updateStatus("🎨 读取成功，正在反色...");

      const newBase64DataUrl = await invertImagePromise(base64);
      const cleanBase64 = newBase64DataUrl.split(",")[1];

      pic.insertInlinePictureFromBase64(cleanBase64, "Replace");

      await context.sync();
      updateStatus("✅ 成功！已反色");
    });
  } catch (e) {
    console.error(e);
    updateStatus("❌ Word 模式失败：" + (e?.message || e));
  }
}

/* =========================
 * 通用：getSelectedDataAsync(Image)
 * ========================= */
function runInvertCommon() {
  Office.context.document.getSelectedDataAsync(
    Office.CoercionType.Image,
    { valueFormat: Office.ValueFormat.Base64 },
    function (asyncResult) {
      if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        updateStatus("❌ 通用读取失败: " + asyncResult.error.message);
        return;
      }

      invertImagePromise(asyncResult.value)
        .then((newBase64DataUrl) => {
          const cleanBase64 = newBase64DataUrl.split(",")[1];

          Office.context.document.setSelectedDataAsync(
            cleanBase64,
            { coercionType: Office.CoercionType.Image },
            (res) => {
              if (res.status === Office.AsyncResultStatus.Failed) {
                updateStatus("❌ 通用替换失败: " + res.error.message);
              } else {
                updateStatus("✅ 成功！已反色（通用模式）");
              }
            }
          );
        })
        .catch((e) => {
          console.error(e);
          updateStatus("❌ 反色计算失败: " + (e?.message || e));
        });
    }
  );
}

/* =========================
 * 图像反色：核心算法
 * 输入：base64（可带/不带 data: 前缀）
 * 输出：data:image/png;base64,xxxx
 * ========================= */
function invertImagePromise(base64Str) {
  return new Promise((resolve, reject) => {
    const img = new Image();

    // 兼容：PPT shape.getImageAsBase64 返回的通常是不带 data: 前缀
    if (base64Str && !base64Str.startsWith("data:")) {
      img.src = "data:image/png;base64," + base64Str;
    } else {
      img.src = base64Str;
    }

    img.onload = () => {
      try {
        const canvas = document.createElement("canvas");
        canvas.width = img.width;
        canvas.height = img.height;

        const ctx = canvas.getContext("2d", { willReadFrequently: true });
        ctx.drawImage(img, 0, 0);

        const imageData = ctx.getImageData(0, 0, canvas.width, canvas.height);
        const data = imageData.data;

        for (let i = 0; i < data.length; i += 4) {
          data[i] = 255 - data[i];         // R
          data[i + 1] = 255 - data[i + 1]; // G
          data[i + 2] = 255 - data[i + 2]; // B
          // Alpha 不变
        }

        ctx.putImageData(imageData, 0, 0);
        resolve(canvas.toDataURL("image/png"));
      } catch (err) {
        reject(err);
      }
    };

    img.onerror = (e) => reject(e);
  });
}

function updateStatus(message) {
  const el = document.getElementById("status");
  if (el) el.innerText = message;
}
