# End User Upload Guide

## Who is this for
This guide is for users who do not run code. You only upload files and collect outputs the next day.

## What you need to do
1. Put your Excel file in the input folder:
   - \\SINSDV030.infineon.com\DEM_AP_MM_SALES\01_SHARING\Digitalisation\AI_Automation\Top_Opportunity\input
2. Wait for nightly processing (runs once every night).
3. Next day, download your output from:
   - \\SINSDV030.infineon.com\DEM_AP_MM_SALES\01_SHARING\Digitalisation\AI_Automation\Top_Opportunity\output

## File naming recommendation
Use clear names so you can identify your output easily.
Example:
- TopOpportunity_MarketX_2026-08-27.xlsx

Your output file name will contain:
- _company_enriched

Example output:
- TopOpportunity_MarketX_2026-08-27_company_enriched.xlsx

## Input file expectations
Your input workbook should contain at least one of these:
1. Website/domain column
2. Contact email column

Optional but helpful:
- Ship-to company name column

## What you get in output
1. Original main sheet remains in place.
2. Added enrichment columns on the right.
3. Data_All Matches tab with detailed match records.

## Important rules
1. Upload only final files for processing.
2. Do not keep the same file open while uploading.
3. Do not upload temporary files (names starting with ~$).
4. If you update a file, re-upload it with a new filename.

## Timing
1. Files uploaded before nightly run are processed the same night.
2. Files uploaded after nightly run are processed the next night.
3. If automation admin enables notifications, you may receive an email when your output file is ready.

## Troubleshooting
1. No output next day:
   - Check if the file was uploaded to the input folder correctly.
   - Check if filename is valid .xlsx.
   - Contact automation support with your filename and upload time.
2. Output exists but looks old:
   - Confirm you uploaded a new filename.
   - Confirm timestamp of the output file.

## Support message template
Use this when raising a support request:
- Input filename:
- Upload date/time:
- Expected output filename:
- Issue observed:
