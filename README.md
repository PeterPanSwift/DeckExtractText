# Keynote Text Extractor

Extract slide text, text hyperlinks, and optional image OCR from Keynote and PowerPoint presentations on macOS.

[English](#english) · [繁體中文](#繁體中文)

[![Platform: macOS](https://img.shields.io/badge/platform-macOS-000000?logo=apple&logoColor=white)](#en-requirements)
[![Language: AppleScript](https://img.shields.io/badge/language-AppleScript-6A5ACD)](#en-project-files)
[![Formats: KEY and PPTX](https://img.shields.io/badge/formats-.key%20%7C%20.pptx-0078D4)](#en-features)

<a id="english"></a>

## English

[切換至繁體中文](#繁體中文)

Keynote Text Extractor reads each slide from Keynote (`.key`) and PowerPoint (`.pptx`) presentations through Keynote, then writes the extracted content to a UTF-8 text file. It supports native slide text, text hyperlinks, tables, grouped objects, skipped slides, and optional OCR through a macOS Shortcut.

<a id="en-features"></a>

### ✨ Features

- Opens `.key` and `.pptx` presentations through Keynote.
- Extracts text from titles, body placeholders, text boxes, shapes, tables, and grouped objects.
- Lists text hyperlinks on their corresponding slides.
- Choose image OCR with the **提取圖片文字** Shortcut, skip images, or export image files without OCR. Image export excludes slide backgrounds.
- Omits empty OCR results, `未辨識到文字`, OCR failures, and image-processing warnings from the output body.
- Updates the output file after every completed slide, so partial progress is retained if processing stops.
- Leaves the source presentation unchanged.

<a id="en-requirements"></a>

### 📋 Requirements

- macOS with Keynote installed.
- Permission for Script Editor or Terminal to control Keynote when macOS requests it.
- For image OCR only: a Shortcut named **提取圖片文字** that accepts an image file and returns recognized text or copies it to the clipboard.

Microsoft PowerPoint and Python are not required to run the extractor. Python 3 is only needed when rebuilding the compiled script after changing the developer source.

<a id="en-quick-start"></a>

### 🚀 Quick start

1. Download [ExtractKeynoteText.scpt](ExtractKeynoteText.scpt).
2. Open it with **Script Editor**.
3. Click **Run**.
4. Select a `.key` or `.pptx` presentation.
5. Select **辨識圖片** (OCR), **跳過圖片** (skip), or **取出圖片** (export image files), then click **繼續**.
6. Choose a new `.txt` output file.

The compiled `.scpt` contains the extraction core. Regular users only need this one file.

### Terminal usage

Run with image OCR:

```sh
osascript ExtractKeynoteText.scpt '/path/to/presentation.key'
```

Read a PowerPoint presentation:

```sh
osascript ExtractKeynoteText.scpt '/path/to/presentation.pptx'
```

Skip all images:

```sh
osascript ExtractKeynoteText.scpt '/path/to/presentation.pptx' --skip-images
```

Choose the output path explicitly:

```sh
osascript ExtractKeynoteText.scpt '/path/to/presentation.key' '/path/to/result.txt'
```

Export images without OCR:

```sh
osascript ExtractKeynoteText.scpt '/path/to/presentation.pptx' '/path/to/result.txt' --export-images
```

This creates `result_圖片/` beside `result.txt`. Files are named by slide and image number, for example `page_003_image_001.png`, automatically converting `.tif` / `.tiff` images to PNG (including uppercase extensions); other exported media formats are retained. Each slide's TXT section lists relative paths such as `[圖片 1] result_圖片/page_003_image_001.png`. Keep the TXT and image folder together when moving them. If the folder name exists, a new suffix is added. Images are saved regardless of whether they contain text; this mode needs no OCR Shortcut and does not access the clipboard. Do not combine `--export-images` with `--skip-images`.

Without an explicit output path, the script creates `<presentation>_文字.txt`. If that name exists, it selects `_文字_2.txt`, `_文字_3.txt`, and so on. An explicitly selected output path must not already exist.

<a id="en-output"></a>

### 📝 Output

The UTF-8 report contains a source path, slide count, processing status, one section per slide, and final statistics.

```text
===== 第 3 頁 =====

【投影片文字】
Course website

[超連結] Course website：https://example.com

[圖片 2]
Text recognized from the second image
```

In OCR mode, image entries appear only when OCR returns useful text. Export mode lists successfully saved image paths instead. The script deliberately excludes the `【圖片 OCR】` heading, no-text responses, OCR errors, and image-processing warnings. Failed image copies or conversions contribute to the warning count; partial image exports remain available if extraction stops.

<a id="en-shortcut"></a>

### 🔍 OCR Shortcut integration

For each referenced image, the script runs:

```sh
shortcuts run '提取圖片文字' --input-path '/path/to/image'
```

It first reads a Shortcut output file and falls back to the clipboard when the Shortcut does not return one. Clipboard contents are saved and restored around each OCR run. Avoid changing the clipboard, editing the presentation, or starting a second extraction while image OCR is running.

The `--skip-images` option does not inspect images, check the Shortcut, run OCR, or access the clipboard.

<a id="en-how-it-works"></a>

### ⚙️ How it works

1. Keynote opens the source presentation and exposes native slide objects through AppleScript.
2. The script exports a temporary PPTX snapshot and unpacks it with macOS `ditto`.
3. It follows `presentation.xml` relationships to preserve slide order and associates text hyperlinks and embedded images with the correct slide.
4. In OCR mode, it processes each slide's referenced images and caches shared media results. In export mode, it copies those images into the output folder; the same media referenced on different slides gets a separate filename for each slide. Duplicate references on one slide are saved once.
5. It writes progress after each slide and removes temporary files after a normal completion.

This snapshot approach avoids unreliable image `file` properties in Keynote's AppleScript interface and distinguishes multiple images that share the same filename.

<a id="en-limitations"></a>

### ⚠️ Scope and limitations

- PowerPoint files are imported by Keynote. Text extraction therefore reflects what Keynote can import from unsupported PowerPoint effects or objects.
- Object order follows Keynote's AppleScript object order and may not match visual reading order.
- Text hyperlinks are listed as `[超連結] label：URL`; duplicate items on the same slide are removed. Click actions attached to images or shapes are excluded.
- OCR and image export use the full exported image and may include cropped or covered content. Native text and image text may overlap. Exported media can include image fills and video preview images; these are not screenshots of the slide's visible image area. Image export excludes images defined as slide backgrounds; an ordinary image placed behind other objects is still a slide object and is retained.
- External linked images are not downloaded. Video and audio are not transcribed.
- Master-layout content that Keynote does not expand into a slide is outside the extraction scope.
- Presenter notes are disabled by default. Developers can set `includePresenterNotes` to `true` in the core and rebuild.

<a id="en-project-files"></a>

### 🧩 Project files

| File | Purpose |
| --- | --- |
| [`ExtractKeynoteText.scpt`](ExtractKeynoteText.scpt) | Ready-to-run compiled script with the core embedded. |
| [`ExtractKeynoteText.applescript`](ExtractKeynoteText.applescript) | Generated, self-contained plain-text launcher. It can also be run directly. |
| [`ExtractKeynoteTextCore.applescript`](ExtractKeynoteTextCore.applescript) | Developer source for extraction, hyperlinks, and OCR. |
| [`build_script.py`](build_script.py) | Embeds the core in the launcher and compiles the `.scpt`. |

After changing the core, rebuild both distributable files:

```sh
python3 build_script.py
```

The current source targets Keynote Creator Studio with bundle ID `com.apple.Keynote`. For the traditional Keynote app, replace it with `com.apple.iWork.Keynote` in the core and rebuild.

<a id="en-validation"></a>

### ✅ Validation

The extractor was exercised with the 29-slide reference Keynote deck and with its PPTX export. The PPTX no-image test completed all 29 slides and recovered 27 text hyperlinks with zero warnings. OCR output still depends on the configured Shortcut and should be reviewed when exact transcription matters.

[Back to the top](#keynote-text-extractor) · [切換至繁體中文](#繁體中文)

---

<a id="繁體中文"></a>

## 繁體中文

[Switch to English](#english)

Keynote Text Extractor 會透過 Keynote 逐頁讀取 Keynote (`.key`) 與 PowerPoint (`.pptx`) 簡報，再將內容寫入 UTF-8 文字檔。它支援原生投影片文字、文字超連結、表格、群組物件、略過的投影片，以及透過 macOS 捷徑執行的選用圖片 OCR。

<a id="zh-features"></a>

### ✨ 功能

- 透過 Keynote 開啟 `.key` 與 `.pptx` 簡報。
- 擷取標題、內文預留位置、文字方塊、形狀、表格與群組物件中的文字。
- 將文字超連結列在所屬投影片中。
- 可選擇透過 **提取圖片文字** 捷徑辨識圖片、跳過圖片，或直接取出圖檔而不執行 OCR。取出圖片時會排除投影片背景。
- 結果內文會省略空白 OCR、`未辨識到文字`、OCR 失敗與圖片處理警告。
- 每完成一頁就更新輸出檔，即使處理中止也能保留已完成的內容。
- 不會更改來源簡報。

<a id="zh-requirements"></a>

### 📋 系統需求

- 已安裝 Keynote 的 macOS。
- macOS 詢問時，允許「腳本編輯器」或「終端機」控制 Keynote。
- 僅圖片 OCR 需要：一個名為 **提取圖片文字** 的捷徑；它必須能接收圖片檔，並直接回傳辨識文字或將文字複製到剪貼簿。

執行擷取不需要 Microsoft PowerPoint 或 Python。只有開發者修改核心後重新建置編譯腳本時才需要 Python 3。

<a id="zh-quick-start"></a>

### 🚀 快速開始

1. 下載 [ExtractKeynoteText.scpt](ExtractKeynoteText.scpt)。
2. 使用「腳本編輯器」開啟。
3. 按下「執行」。
4. 選擇 `.key` 或 `.pptx` 簡報。
5. 從清單選擇「辨識圖片」、「跳過圖片」或「取出圖片」，再按「繼續」。
6. 選擇一個新的 `.txt` 輸出檔案。

編譯完成的 `.scpt` 已內嵌擷取核心，一般使用者只需要這一個檔案。

### 終端機操作

執行圖片 OCR：

```sh
osascript ExtractKeynoteText.scpt '/簡報路徑/課程.key'
```

讀取 PowerPoint 簡報：

```sh
osascript ExtractKeynoteText.scpt '/簡報路徑/課程.pptx'
```

跳過所有圖片：

```sh
osascript ExtractKeynoteText.scpt '/簡報路徑/課程.pptx' --skip-images
```

指定輸出路徑：

```sh
osascript ExtractKeynoteText.scpt '/簡報路徑/課程.key' '/輸出路徑/結果.txt'
```

取出圖片而不執行 OCR：

```sh
osascript ExtractKeynoteText.scpt '/簡報路徑/課程.pptx' '/輸出路徑/結果.txt' --export-images
```

這會在 `結果.txt` 旁建立 `結果_圖片/` 資料夾。圖檔依投影片頁數與圖片序號命名，例如 `page_003_image_001.png`，遇到 `.tif`／`.tiff`（含大寫副檔名）會自動轉成 PNG，其他格式則保留匯出媒體的格式。TXT 每頁會列出相對路徑，例如 `[圖片 1] 結果_圖片/page_003_image_001.png`；移動結果時請一起搬移 TXT 與圖片資料夾。資料夾名稱若已存在，會加上新的數字尾碼。此模式不論圖片是否有文字都會存檔，不需要 OCR 捷徑，也不存取剪貼簿。`--export-images` 不可與 `--skip-images` 同時使用。

未指定輸出路徑時，腳本會建立 `<簡報名稱>_文字.txt`。如果檔名已存在，會依序改用 `_文字_2.txt`、`_文字_3.txt`。若明確指定輸出路徑，該檔案不可已經存在。

<a id="zh-output"></a>

### 📝 輸出內容

UTF-8 報告會包含來源路徑、投影片頁數、處理狀態、每頁內容與最後的擷取統計。

```text
===== 第 3 頁 =====

【投影片文字】
課程網站

[超連結] 課程網站：https://example.com

[圖片 2]
從第二張圖片辨識出的文字
```

辨識模式只有在 OCR 回傳有效文字時才會出現圖片項目；取出圖片模式則列出成功儲存的圖檔路徑。腳本會刻意排除 `【圖片 OCR】` 標題、無文字回覆、OCR 錯誤與圖片處理警告。圖檔複製或轉換失敗會累計警告次數；處理中止時仍保留已取出的圖片。

<a id="zh-shortcut"></a>

### 🔍 OCR 捷徑整合

腳本會對每張被投影片引用的圖片執行：

```sh
shortcuts run '提取圖片文字' --input-path '/圖片路徑/image.png'
```

腳本會優先讀取捷徑輸出檔；捷徑沒有直接回傳檔案時，則讀取剪貼簿。每次 OCR 前後都會保存並還原剪貼簿。辨識圖片期間請避免變更剪貼簿、編輯簡報，或同時開始另一個擷取程序。

使用 `--skip-images` 時，不會檢查圖片、檢查捷徑、執行 OCR 或存取剪貼簿。

<a id="zh-how-it-works"></a>

### ⚙️ 運作方式

1. Keynote 開啟來源簡報，腳本透過 AppleScript 讀取原生投影片物件。
2. 腳本匯出暫存 PPTX，並使用 macOS 內建的 `ditto` 解開。
3. 它依照 `presentation.xml` 關聯保留正確頁序，並將文字超連結與內嵌圖片配回所屬投影片。
4. 啟用 OCR 時，腳本會處理每頁引用的圖片，並快取跨頁共用的媒體辨識結果。取出圖片模式則將圖片複製到輸出資料夾；跨頁共用圖片會依各頁另存不同檔名，同頁重複引用同一媒體只存一次。
5. 每完成一頁就寫入進度；正常完成後會移除暫存檔案。

使用暫存 PPTX 是為了避開 Keynote AppleScript 介面中不穩定的圖片 `file` 屬性，也能區分原始檔名相同的多張圖片。

<a id="zh-limitations"></a>

### ⚠️ 範圍與限制

- PowerPoint 檔案會先由 Keynote 匯入；若內容含 Keynote 不支援的 PowerPoint 效果或物件，擷取結果以 Keynote 能匯入的內容為準。
- 物件順序依 Keynote 的 AppleScript 物件順序排列，可能與畫面閱讀順序不同。
- 文字超連結以 `[超連結] 顯示文字：URL` 顯示，同頁重複項目會去除；圖片或形狀本身的點擊動作不列入。
- OCR 與取出圖片使用完整的匯出圖片，因此可能包含被裁切或遮住的內容；原生文字與圖片文字也可能重複。匯出媒體可能包含形狀圖片填滿及影片預覽圖，並非投影片中可見圖片區域的截圖。取出圖片會排除設定為投影片背景的圖片；若一般圖片只是被放在其他物件後方，仍會視為頁面物件保留。
- 不會下載外部連結圖片，也不會轉錄影片或音訊。
- Keynote 未展開至個別投影片的母片或版面內容不在擷取範圍。
- 講者備忘稿預設關閉。開發者可在核心將 `includePresenterNotes` 設為 `true` 後重新建置。

<a id="zh-project-files"></a>

### 🧩 專案檔案

| 檔案 | 用途 |
| --- | --- |
| [`ExtractKeynoteText.scpt`](ExtractKeynoteText.scpt) | 可直接執行的編譯腳本，已內嵌核心。 |
| [`ExtractKeynoteText.applescript`](ExtractKeynoteText.applescript) | 自動產生且可獨立執行的純文字啟動器。 |
| [`ExtractKeynoteTextCore.applescript`](ExtractKeynoteTextCore.applescript) | 負責擷取、超連結與 OCR 的開發者核心原始碼。 |
| [`build_script.py`](build_script.py) | 將核心嵌入啟動器，並編譯 `.scpt`。 |

修改核心後，請重新建立兩個發布檔案：

```sh
python3 build_script.py
```

目前原始碼使用 Keynote Creator Studio 的 bundle ID `com.apple.Keynote`。若使用傳統 Keynote，請在核心將它改成 `com.apple.iWork.Keynote` 後重新建置。

<a id="zh-validation"></a>

### ✅ 驗證

本工具已使用 29 頁的參考 Keynote 簡報及其 PPTX 匯出檔測試。PPTX 跳過圖片測試完成全部 29 頁，擷取 27 個文字超連結，警告為 0。OCR 結果仍取決於使用者設定的捷徑；需要精確逐字稿時應再人工校對。

[回到頂端](#keynote-text-extractor) · [Switch to English](#english)
