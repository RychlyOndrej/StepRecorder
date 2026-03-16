# StepRecorder

`StepRecorder` is a Windows desktop app (WPF, `.NET 9`) for recording step-by-step workflows using screenshots and exporting them as documentation.

## Features

- step capture by:
  - mouse click
  - hotkeys (full screen / crop around cursor)
- optional capture limited to a selected application window
- automatic step numbering
- per-step metadata:
  - timestamp
  - window/process name
  - pressed keys
- screenshot annotations:
  - click highlight (color + radius)
  - optional step number badge on image
- double-click protection (avoids duplicate captures)
- safe unique session folders (no overwrite on same name)

## Export formats

- `PDF`
- `MHT`
- `Word (.docx)`

After export, the app shows generated files and lets you open them directly.

## Requirements

- Windows
- `.NET 9 SDK`
- Visual Studio 2022 (recommended)

Used packages:

- `QuestPDF`
- `DocumentFormat.OpenXml`

## Run in Visual Studio

1. Open the solution/repository in Visual Studio.
2. Set `StepRecorder` as startup project.
3. Run (`F5`).

## Quick usage

1. In **Session** tab, set name, output folder, and export formats.
2. In **Recording** tab, configure capture behavior and annotations.
3. Click **Start recording**.
4. Perform steps in the target app.
5. Click **Stop and export**.
6. In **Export** tab, run export and open output files.

## Notes

- For cursor crop in “selected window only” mode, the crop rectangle is shifted to stay inside the app window while keeping fixed crop size.
- If a session folder with the same name already exists, the app creates `Name (2)`, `Name (3)`, etc.

## Project structure (simplified)

- `StepRecorder/MainWindow.xaml` – main UI
- `StepRecorder/MainWindow.xaml.cs` – UI logic and session control
- `StepRecorder/Services/RecordingService.cs` – recording orchestration
- `StepRecorder/Services/ScreenCaptureService.cs` – screenshots and crop logic
- `StepRecorder/Services/ExportService.cs` – export to PDF/MHT/DOCX
- `StepRecorder/Models/*` – data models and settings

---

# StepRecorder (CZ)

`StepRecorder` je desktopová aplikace pro Windows (WPF, `.NET 9`), která zaznamenává pracovní postup krok za krokem pomocí screenshotù a následnì ho umí exportovat do dokumentu.

## Co aplikace umí

- nahrávání krokù pøi:
  - kliknutí myší
  - klávesovıch zkratkách (celá obrazovka / vıøez kolem myši)
- omezení záznamu na vybrané okno aplikace
- automatické èíslování krokù
- metadata ke krokùm:
  - èas
  - název okna/proces
  - stisknuté klávesy
- anotace screenshotù:
  - zvıraznìní kliknutí (barva + velikost)
  - volitelnì èíslo kroku na obrázku
- ochrana proti duplicitnímu záznamu pøi dvojkliku
- bezpeèné vytváøení unikátních sloek relací (bez pøepisování pøi stejném názvu)

## Export

Aplikace podporuje export do:

- `PDF`
- `MHT`
- `Word (.docx)`

Po exportu aplikace zobrazí seznam vytvoøenıch souborù a umoní je pøímo otevøít.

## Poadavky

- Windows
- `.NET 9 SDK`
- Visual Studio 2022 (doporuèeno)

Pouité balíèky:

- `QuestPDF`
- `DocumentFormat.OpenXml`

## Rychlé pouití

1. Na kartì **Relace** nastavte název, vıstupní sloku a formát exportu.
2. Na kartì **Nahrávání** nastavte zpùsob snímání a chování anotací.
3. Kliknìte na **Spustit nahrávání**.
4. Proveïte kroky v cílové aplikaci.
5. Kliknìte na **Zastavit a exportovat**.
6. Na kartì **Export** spuste export a otevøete vısledné soubory.

## Poznámky

- U vıøezu kolem myši se pøi reimu „pouze vybrané okno“ vıøez posouvá tak, aby drel pevnou velikost a nebral okolí mimo aplikaci.
- Pokud u existuje sloka se stejnım názvem relace, vytvoøí se automaticky varianta `Název (2)`, `Název (3)` atd.

## Struktura projektu (zjednodušenì)

- `StepRecorder/MainWindow.xaml` – hlavní UI
- `StepRecorder/MainWindow.xaml.cs` – logika UI a ovládání relace
- `StepRecorder/Services/RecordingService.cs` – orchestrace záznamu
- `StepRecorder/Services/ScreenCaptureService.cs` – screenshoty a crop
- `StepRecorder/Services/ExportService.cs` – export do PDF/MHT/DOCX
- `StepRecorder/Models/*` – datové modely a nastavení