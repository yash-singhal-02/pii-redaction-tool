# PII Redaction Tool

## Overview

This project redacts personally identifiable information from the supplied Red Herring Prospectus and creates a new DOCX file. I used Python because the document is a DOCX and the approach can be kept simple and easy to maintain.

## Approach

The solution uses a small hybrid approach instead of a large NLP framework:

- `python-docx` is used to read and write the document.
- Regular expressions handle structured PII such as email addresses, phone numbers, SSNs, credit cards, dates of birth and IPv4 addresses.
- Credit-card candidates are checked with the Luhn algorithm to reduce false positives.
- Names and company names use document context such as `Contact Person`, promoter/director fields and company suffixes such as `Limited`, `Private Limited` and `LLP`.
- Addresses use address keywords, Indian PIN-code patterns and table-column context.
- Tables, headers and footers are included in the text scan.
- The supplied document also contains identity-card images. I added OCR using `pytesseract` so that embedded Aadhaar/PAN images are detected. For these two card layouts, the personal fields, photographs, QR codes and identity numbers are covered with `[REDACTED]` boxes before the image is put back into the DOCX.

The same detected text value gets the same fake replacement throughout the document. For image-based ID cards, I use `[REDACTED]` instead of inventing a fake government ID, because putting a fake ID number into a government-card image could be misleading.

## PII types supported

1. Full names
2. Email addresses
3. Phone numbers
4. Company names
5. Physical/mailing addresses
6. SSNs
7. Credit card numbers
8. Dates of birth
9. IP addresses
10. PII contained in the supplied Aadhaar/PAN images

## How to run

Install the dependencies:

```bash
pip install -r requirements.txt
```

Tesseract OCR must also be installed and available on the system PATH.

Run the redactor:

```bash
python pii_redactor.py "Red Herring Prospectus(1).docx" "redacted_output.docx" --report run_report.json
```

Run the evaluation:

```bash
python evaluate.py "Red Herring Prospectus(1).docx" "redacted_output.docx"
```

## Evaluation approach

I used two levels of evaluation.

### 1. Synthetic text test set

I created 45 small test cases:

- 27 positive cases: 3 examples for each of the 9 required PII types.
- 18 negative cases containing values that look similar to PII but should not be redacted, such as order numbers, ordinary dates and invalid IP addresses.

A positive case is counted as correct when the expected PII type is detected. A negative case is counted as correct when nothing is detected. From these counts I calculate accuracy, precision and recall.

Final result:

- Accuracy: **100.00%**
- Precision: **100.00%**
- Recall: **100.00%**
- True positives: 27
- False positives: 0
- True negatives: 18
- False negatives: 0

Each individual PII type also achieved 3/3 recall in this small test set.

### 2. Supplied document validation

I also checked the generated DOCX against the actual supplied prospectus. The document contains text PII as well as embedded identity-card images. The final run checked:

- 26 unique source email values — 0 remained
- 30 manually audited person names — 0 remained
- 10 manually audited company names — 0 remained
- 5 manually audited address examples — 0 remained
- 2 identity-card images — both were redacted
- 14 manually defined PII regions across those two images — 14/14 were redacted

The 100% precision/recall figures above are for the controlled 45-case evaluation set. They should not be interpreted as proof that the tool will be perfect on every unseen document.

## Trade-offs and limitations

This is intentionally a small student-friendly implementation. A production system could use Microsoft Presidio, spaCy NER or another trained model to improve detection of names and less structured addresses.

Regex works well for structured values but can miss unusual formats. Name and company detection depends on context, so an isolated name without useful surrounding text can be missed. OCR can also make mistakes on low-quality images. The identity-card image handling therefore combines OCR-based card identification with layout-based redaction for the actual images present in the supplied prospectus.

The supplied prospectus does not provide normal text examples of SSNs or credit-card numbers, so those were evaluated using synthetic test cases. The Aadhaar and PAN information was present inside images, which is why image processing was added after inspecting the complete DOCX rather than relying only on extracted text.

## Author
Yash Singhal
