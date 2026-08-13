# Evaluation Report

## 1. Objective

The objective was to check whether the redaction tool can identify the required PII types while avoiding obvious false positives, and to verify that PII contained in the supplied document's embedded identity-card images is also removed.

## 2. Text evaluation dataset

The evaluation uses 45 controlled cases:

| Category | Cases |
|---|---:|
| Full names | 3 positive |
| Email addresses | 3 positive |
| Phone numbers | 3 positive |
| Company names | 3 positive |
| Physical addresses | 3 positive |
| SSNs | 3 positive |
| Credit cards | 3 positive |
| Dates of birth | 3 positive |
| IP addresses | 3 positive |
| Hard negatives | 18 |
| **Total** | **45** |

For each positive case, the expected PII type and value are known. For each negative case, the expected result is no detection.

## 3. Metrics

The final run produced:

| Metric | Result |
|---|---:|
| True positives | 27 |
| False positives | 0 |
| True negatives | 18 |
| False negatives | 0 |
| Accuracy | **100.00%** |
| Precision | **100.00%** |
| Recall | **100.00%** |

Formulas used:

```text
Accuracy  = (TP + TN) / (TP + TN + FP + FN)
Precision = TP / (TP + FP)
Recall    = TP / (TP + FN)
```

All nine required PII categories achieved 3/3 recall in the synthetic test set.

## 4. Image evaluation

The supplied prospectus contains two identity-card images that were not represented completely in the DOCX text extraction. One is an Aadhaar image and one is a PAN image.

For image evaluation, I defined 14 PII regions across those two images. These regions cover personal photos, names, father's names, dates of birth, identity numbers, QR codes, signatures and addresses.

Result:

- Identity-card images checked: **2**
- PII regions checked: **14**
- PII regions redacted: **14**
- Region redaction rate: **100.00%**

The image check does not expose the original PII values in this report. It checks whether the known sensitive regions are covered after processing.

## 5. Validation against the supplied prospectus

A second validation was performed on the generated DOCX:

| Check | Found/Audited | Remaining in output |
|---|---:|---:|
| Unique source emails | 26 | 0 |
| Audited person names | 30 | 0 |
| Audited company names | 10 | 0 |
| Audited addresses | 5 | 0 |
| Identity-card images | 2 | 0 unredacted identity images |

## 6. Interpretation

The synthetic evaluation gives 100% accuracy, precision and recall, but the test set is small. It is mainly intended to demonstrate that the implemented rules work for the selected formats and that hard-negative examples are not unnecessarily redacted.

The document-level validation is separate because some required PII types, such as SSNs and credit-card numbers, are not present as normal text in the supplied prospectus. The Aadhaar and PAN data are image-based, so they are evaluated through the image-redaction checks instead.

## 7. Limitations

- Regex patterns do not cover every possible formatting variation.
- Name/company detection works best when the document provides context.
- OCR quality depends on the image quality and language.
- The Aadhaar/PAN image redaction uses a small layout rule for the card images present in this assignment; a production solution should use a more general document/image PII detector.
- The 100% metrics should not be treated as a guarantee for unseen documents.
