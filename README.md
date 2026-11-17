# First Steps with Office Scripts in Excel

Give your mouse a vacation. In this 45-minute webinar, turn a repetitive time-report formatting task into an Office Script you can inspect, test, and run again. See where AI helps you draft or explain the code, and where a saved script is the better way to repeat an agreed set of steps.

No coding experience is required. Watching is welcome; you do not need to type along. The recording will be made available after the webinar.

## Before you begin

Download this repository, open a copy of `record-actions.xlsx` in Excel, and look for the **Automate** tab. Office Scripts is supported in Excel for the web, supported Windows builds, and Mac; access depends on your account, license, and administrator settings. See Microsoft's [current requirements](https://learn.microsoft.com/en-us/office/dev/scripts/testing/platform-limits). Use OneDrive or SharePoint for the sharing demonstration. If Automate is unavailable, watch the demo and try the files later. An AI subscription is not required for this session.

The workbooks contain fictional people and tasks. Dates start as Excel serial numbers on purpose: displaying them as dates is part of the exercise. Work on copies so that you can reopen the originals to reset.

## Files

| File | Use |
| --- | --- |
| `first-steps-office-scripts.pptx` | Slides and timed presenter notes; 45 minutes including questions. |
| `record-actions.xlsx` | Record on `report` (six rows), then test on `report-2` (seven rows). |
| `format-time-report.ts` | Copy-ready final script with a named sheet, flexible row count, and threshold input. |
| `share-office-scripts.xlsx` | A separate copy of the seven-row report for the sharing demonstration. |
| `office-scripts-vs-vba.pdf` | Two-page reference: AI, Office Scripts, VBA, and the demo checks. |

## The 45-minute route

| Minutes | Focus |
| --- | --- |
| 0-5 | Choose the tool, check access, and preview the result. |
| 5-20 | Record formatting, rerun it, and spot what a fixed range misses. |
| 20-30 | Read the code, ask AI for a focused improvement, and review the final script. |
| 30-35 | Test the extra row, a second run, and a different threshold. |
| 35-40 | Share a reviewed script and decide when AI, a script, or both fit. |
| 40-45 | RECORD, CODE, AI, or SHARE: revisit a step and take questions. |

## Record, run, and improve

1. In `record-actions.xlsx`, activate `report`. Select **Automate > Record Actions** and keep **absolute references** for this demonstration. Format **A2:A7** as dates, autofit **A:D**, and apply a conditional-formatting rule to **D2:D7** for values **greater than 3**. Stop recording and name the script `Format time report - recorded`.
2. Open the recorded code. Confirm that it targets the active worksheet and the fixed ranges above. If it names `report` explicitly, change only that sheet selection to `workbook.getActiveWorksheet()` for this transfer experiment; recorder output can vary. Activate `report-2` and run it. D8 contains **4.5**, but a rule applied only to D2:D7 misses it. A8 also remains a serial number when only A2:A7 receives date formatting. The recorder captures your choices; it does not infer how next week's report might grow.
3. Use the AI prompt below to discuss a targeted improvement. AI can propose code and explain it; the proposal still needs review. The supplied final script is ready if AI is unavailable or its answer differs.
4. Choose **Automate > New Script > Create in Code Editor**, replace the starter code with the contents of `format-time-report.ts`, and save it as `Format time report`. Run it with `sheetName` = `report-2` and `threshold` = `3`. You paste the text into the editor; the `.ts` file is not an embedded workbook macro.
5. Check the results below, then open `share-office-scripts.xlsx` from OneDrive or SharePoint and run the same saved script with its defaults. In the script details, use **Associate with workbook** and **Add button to worksheet** if your account permits it. Exact labels can vary. Native script sharing is within your organization and subject to workbook and script permissions. For someone outside it, share the `.ts` source and sample workbook so they can review and save their own script; that is an independent copy, not a live link to yours.

## Where AI fits

Use AI when you need help turning an idea into steps, explaining an unfamiliar line, or troubleshooting a specific error. Use a saved Office Script when the rules are settled and you want to repeat them on similarly structured workbooks. Often the useful combination is **AI drafts or explains; you review and test; the script repeats the accepted steps**. Some AI tools can also edit workbooks directly; the distinction here is whether you want a fresh interpretation of a request or an explicit routine you maintain. Office Scripts does not require AI.

Paste this prompt into an approved AI assistant with the recorded script. It uses the sample workbook's structure, so no real business data is needed:

```text
Help me improve this Excel Office Script, using the ExcelScript API (not VBA or Office.js).
First explain which worksheet and ranges it changes.
The report starts at A1 with Date, Name, Task, Hours Spent in A1:D1.
Dates are numeric Excel dates and hours are nonnegative numbers. There are no blank data rows.
Use a sheetName parameter defaulting to "report-2" and a numeric threshold defaulting to 3.
Find all populated rows in A:D. Format the dates, autofit the columns, and highlight
hours strictly greater than the threshold with red fill (#CF3338) and white text.
Keep all cell values unchanged. Check the headers before making changes.
The script owns the hours cells' conditional-formatting rules: replace them so reruns
do not add duplicates. Explain that side effect. Do not modify other worksheets.
Explain the changes and give tests for a seventh row, a repeated run, and a threshold of 4.

[Paste the recorded script here.]
```

For an error, share the exact message and the relevant lines: "Explain this Office Scripts error and propose the smallest fix. Tell me what to retest." Keep reviewing the sheet, range, condition, and side effects yourself; a plausible explanation is not proof of correct code.

## Check the result

| Run | Expected result |
| --- | --- |
| `report`, threshold 3 | Six data rows; D4 = 3.5 and D6 = 4 highlighted. |
| `report-2`, threshold 3 | Seven data rows; D4 = 3.5, D6 = 4, D8 = 4.5 highlighted. All dates formatted. |
| `report-2`, threshold 3, again | Same result; one conditional-formatting rule covering D2:D8. |
| `report-2`, threshold 4 | Only D8 = 4.5 highlighted. Exactly 4 does not pass "greater than 4." |
| Wrong sheet name or changed headers | An explanatory error before formatting is changed. |

The console reports the number of rows and values above the threshold. The script changes formatting, not data. It replaces existing conditional-formatting rules in the used hours range; adapt that choice before using a workbook with rules you need to keep. The report must start at A1, have the four exact headers, contain numeric dates/hours, and have no blank data rows or other content in columns A:D. Conditional formatting responds to changed values within its range; rerun the script to include newly appended rows.

**One next action:** record one small task on a copy of your own workbook, then test it on a second example before sharing it.

## Further reading

- George's [guide to Record Actions](https://stringfestanalytics.com/how-to-understand-record-actions-in-office-scripts-for-excel/) and [guide to main()](https://stringfestanalytics.com/how-to-understand-function-main-in-office-scripts-for-excel/).
- Microsoft: [Office Scripts and VBA](https://learn.microsoft.com/en-us/office/dev/scripts/resources/vba-differences), [conditional-formatting examples](https://learn.microsoft.com/en-us/office/dev/scripts/resources/samples/conditional-formatting-samples), and [sharing Office Scripts](https://support.microsoft.com/en-us/excel/office-scripts/sharing-office-scripts-in-excel).

Materials revised November 17, 2026. Features and availability can vary by account and update channel.
