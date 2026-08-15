# Generate an Excel Worksheet Using XCO XLSX in SAP BTP ABAP Environment

This guide shows how to generate an Excel worksheet programmatically using the **XCO XLSX library** in the **SAP BTP ABAP Environment**.

## Prerequisites

* SAP BTP ABAP Environment system
* ABAP Development Tools (ADT)
* An ABAP class or executable implementation where the XCO XLSX API can be used
* Basic knowledge of ABAP syntax

## 1. Create an Empty XLSX Document

Use the XCO XLSX API to create an empty workbook:

```abap
DATA(lo_document) = xco_cp_xlsx=>document->empty( ).
```

The document contains a workbook with an initial worksheet.

## 2. Get the Worksheet

Access the first worksheet by its position:

```abap
DATA(lo_worksheet) =
  lo_document->workbook->worksheet->at_position( 1 ).
```

You can then use the worksheet object to write cells, ranges, or tables.

## 3. Write Data to the Worksheet

For example, write a small table to cells `A1:C3`:

```abap
lo_worksheet->select(
  xco_cp_xlsx=>selection=>pattern( 'A1:C3' )
)->row_stream(
)->operation->write_from(
  VALUE #(
    ( 'ID' 'Name' 'Amount' )
    ( '1'  'Test' '100' )
    ( '2'  'Demo' '200' )
  )
)->execute( ).
```

The resulting worksheet contains:

| ID | Name | Amount |
| -- | ---- | -----: |
| 1  | Test |    100 |
| 2  | Demo |    200 |

## 4. Serialize the XLSX Document

After writing the worksheet data, serialize the document:

```abap
DATA(lv_file_content) = lo_document->write_to( ).
```

`lv_file_content` contains the generated Excel workbook as an `xstring`.

This value can be passed to an HTTP/OData response or another file-handling mechanism.

## 5. Complete Example

A simple method can look like this:

```abap
METHOD generate_excel.

  DATA(lo_document) = xco_cp_xlsx=>document->empty( ).

  DATA(lo_worksheet) =
    lo_document->workbook->worksheet->at_position( 1 ).

  lo_worksheet->select(
    xco_cp_xlsx=>selection=>pattern( 'A1:C3' )
  )->row_stream(
  )->operation->write_from(
    VALUE #(
      ( 'ID' 'Name' 'Amount' )
      ( '1'  'Test' '100' )
      ( '2'  'Demo' '200' )
    )
  )->execute( ).

  DATA(lv_file_content) = lo_document->write_to( ).

ENDMETHOD.
```

## 6. Returning the File from BTP ABAP

In the **SAP BTP ABAP Environment**, ABAP code does not directly write a file to the user's local PC filesystem.

The recommended approach is:

```text
ABAP Method
    |
    v
XCO XLSX Document
    |
    v
write_to( )
    |
    v
XSTRING
    |
    v
OData / HTTP Response
    |
    v
Browser Download
    |
    v
Local .xlsx File
```

For a browser download, the response should use the appropriate Excel MIME type:

```text
application/vnd.openxmlformats-officedocument.spreadsheetml.sheet
```

The `Content-Disposition` header can be set similar to:

```text
attachment; filename="MyFile.xlsx"
```

## 7. Working with Dynamic Data

For real business scenarios, replace the hard-coded values with an internal table.

For example:

```abap
DATA lt_data TYPE TABLE OF your_structure.

" Fill lt_data ...

lo_worksheet->select(
  xco_cp_xlsx=>selection=>pattern( 'A1:C100' )
)->row_stream(
)->operation->write_from(
  REF #( lt_data )
)->execute( ).
```

The exact range and data structure should match the number and type of columns being written.

## 8. Important Notes

### `xstring`

The result of:

```abap
lo_document->write_to( )
```

is binary XLSX content represented as an ABAP `xstring`.

Do not treat this content as ordinary text.

### Local File System

BTP ABAP Environment does not provide the traditional ABAP application-server file handling model for writing directly to your local computer.

For example, do not expect:

```abap
OPEN DATASET ...
```

to save a file directly on your laptop.

Instead, expose the generated `xstring` through an appropriate service/API and let the client download it.

### File Extension

Use:

```text
.xlsx
```

for an XLSX workbook.

### MIME Type

Use:

```text
application/vnd.openxmlformats-officedocument.spreadsheetml.sheet
```

for an XLSX download.

## 9. Recommended Architecture

For a production BTP ABAP application, a common architecture is:

```text
RAP / OData Service
        |
        v
   ABAP Action
        |
        v
   XCO XLSX API
        |
        v
  Generated XSTRING
        |
        v
 Media / HTTP Response
        |
        v
 Browser Download
```

This keeps Excel generation on the ABAP backend while allowing the end user to download the generated workbook.

## References

* SAP XCO XLSX API documentation
* SAP BTP ABAP Environment documentation
* SAP ABAP RESTful Application Programming Model (RAP) documentation
