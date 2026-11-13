Realistic PDF -> Word (.docx) test set
=====================================

Ten PDFs made the way real users' PDFs are made, by two producers:

  * Microsoft Word 2016 "Save as PDF" (ExportAsFixedFormat), built through the Word object
    model. The .docx Word saved next to each PDF is the ground truth for the conversion.
  * Google Chrome print-to-PDF (Skia/PDF), via puppeteer page.pdf({format:'A4',
    printBackground:true}). The .html next to each PDF is the source.

All text is real, selectable text (no scans). No lorem ipsum. Names, companies, addresses,
phone numbers and bank details are fictional.

Generators (re-run to rebuild): gen\pics.ps1 (draws the PNG pictures), gen\word.ps1
(-Name resume|report|letter|newsletter|paper|invoice_word -Out <dir> -Pics <dir> -Tmp <dir>),
gen\web_samples.js (node, needs puppeteer-core and Chrome; args: <outDir> <picsDir>).
COM trap: pass plain strings to Word methods. A Join-Path result handed to
ExportAsFixedFormat makes hidden Word hang forever.


File               Pages  Producer   Layout features exercised
-----------------  -----  ---------  -------------------------------------------------------------
resume.pdf           1    Word       US Letter. Calibri / Calibri Light. 26pt coloured name, coloured
 (+ resume.docx)                     subtitle, small grey contact line with | separators. Body is a
                                     borderless 1x2 table: shaded left sidebar (Skills with typed
                                     bullet characters, Education, Certifications, Languages, small
                                     ALL-CAPS letter-spaced headings) and a main column (justified
                                     profile, job title + em dash + company on one line, grey italic
                                     dates, "List Bullet" bullets with hanging indent, bold run-in
                                     labels).

report.pdf           3    Word       A4. Title + Subtitle styles, Heading 1/2 (Calibri Light, blue),
 (+ report.docx)                     justified body, mixed bold runs, superscript footnote reference
                                     and footnote with separator line, two "List Number" lists (the
                                     second restarts at 1), "List Bullet" lists, caption-style table
                                     title, 4-column table with dark shaded header row (white bold
                                     text), right-aligned numeric columns, banded rows, wrapped cells,
                                     centred bar-chart picture with caption, right-aligned running
                                     header, centred footer "Confidential | Page X of Y" (fields).

letter.pdf           1    Word       A4 business letter. Georgia. Letterhead as a borderless 1x2
 (+ letter.docx)                     table: round logo picture left, right-aligned company name and
                                     grey contact lines, thick green rule under it. Date, 6-line
                                     address block (no spacing), grey reference line, bold subject,
                                     justified body, bullet list, "Yours sincerely", transparent PNG
                                     signature picture, name/title block, italic enclosure line, small
                                     centred grey legal footer.

newsletter.pdf       1    Word       A4. Full-width title banner = paragraph shading (plum) with
 (+ newsletter.docx)                 white letter-spaced Georgia 30pt + issue line; continuous section
                                     break into 2 text columns (22pt gap), Georgia headings, justified
                                     Segoe UI body, inline picture with italic caption inside a column,
                                     pull quote with gold top/bottom paragraph borders and large
                                     italic coloured text, attribution line, bullets, "diary" lines
                                     with coloured bold dates; the second column is shorter
                                     (unbalanced columns).

paper.pdf            2    Word       US Letter academic paper. Times New Roman 11pt throughout.
 (+ paper.docx)                      Centred 16pt bold two-line title, authors with superscript
                                     affiliation marks, italic affiliation line, centred "Abstract",
                                     abstract indented both sides and justified at 10pt, bold/italic
                                     keywords line, numbered headings typed as text ("1  Introduction",
                                     "2.1  Study corridors", italic level 2), justified paragraphs
                                     with first-line indent, footnote, centred table title above a
                                     booktabs-style 5-column table (rules only above/below header and
                                     at the end, right-aligned numbers, wrapped header cell),
                                     references with hanging indent, centred page number footer.

invoice_word.pdf     1    Word       A4. Arial. Two-column header (company block left, big teal
 (+ invoice_word.docx)               "INVOICE" + right-aligned invoice no./dates), Bill-to / Project
                                     blocks side by side, 5-column item table with dark header, thin
                                     grey row rules only, right-aligned Qty/price/amount columns,
                                     right-aligned totals table (Subtotal / Discount / VAT / shaded
                                     "Total due" row), payment details, notes in grey, centred italic
                                     thank-you line, small grey footer.

invoice_web.pdf      1    Chrome     A4, no page margin. Segoe UI. Full-bleed dark blue header band
 (+ invoice_web.html)                with logo image + company name, light "INVOICE" right; three
                                     rounded bordered info boxes in a row (flexbox); item table with
                                     tinted header and grey sub-description lines under each item;
                                     right-aligned money; "How to pay" notes next to a totals box with
                                     filled "Amount due" row; footer pinned to page bottom with
                                     left and right text.

brochure.pdf         2    Chrome     A4. Arial + Georgia. Full-bleed gradient hero band with kicker
 (+ brochure.html)                   (letter-spaced caps), 32pt serif headline over two lines,
                                     paragraph, pill button; three feature cards in a row (top border,
                                     numbered circle icons, rounded corners, tinted fill); dark stats
                                     strip with four big numbers; CSS 2-column justified section that
                                     breaks across pages; footer rule with two-sided contact line.

form_web.pdf         1    Chrome     A4 filled application form. Arial labels, Courier New "typed"
 (+ form_web.html)                   answers in blue. Dark section header bars, labelled boxes in
                                     rows with cell borders (field grid), checkboxes drawn as small
                                     bordered squares, some ticked with a check mark glyph, inline
                                     Yes/No choice, multi-line free-text box, signature in Segoe
                                     Script, dashed grey "office use only" section with empty boxes.

cv_web.pdf           1    Chrome     A4 modern CV. Segoe UI (Light/Semibold/Bold). Full-height dark
 (+ cv_web.html)                     left sidebar (72mm) with initials circle, Contact, Skills with
                                     drawn progress bars, Education, Languages bullets, Interests;
                                     main column with 26pt name, coloured role, letter-spaced
                                     section headings with coloured underline rule, job title
                                     left + dates right on one line, italic company line, bullets,
                                     justified summary.
