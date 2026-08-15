CLASS zcl_excel_create DEFINITION
  PUBLIC FINAL
  CREATE PUBLIC.

  PUBLIC SECTION.
    TYPES: BEGIN OF lty_template,
             space TYPE c LENGTH 100,
             id1   TYPE c LENGTH 2000,
             Id2   TYPE c LENGTH 200,
             id3   TYPE c LENGTH 200,
             id4   TYPE c LENGTH 200,
             id5   TYPE c LENGTH 600,
             id6   TYPE c LENGTH 40,
             id7   TYPE c LENGTH 40,
             id8   TYPE c LENGTH 40,
             id9   TYPE c LENGTH 40,
             id10  TYPE c LENGTH 40,
             id11  TYPE c LENGTH 40,
             id12  TYPE c LENGTH 40,
             id13  TYPE c LENGTH 40,
             id14  TYPE c LENGTH 40,
             id15  TYPE c LENGTH 40,
             id16  TYPE c LENGTH 40,
             id17  TYPE c LENGTH 40,
             id18  TYPE c LENGTH 40,
             id19  TYPE c LENGTH 40,
             id20  TYPE c LENGTH 40,
             id21  TYPE c LENGTH 40,
             id22  TYPE c LENGTH 40,
             id23  TYPE c LENGTH 40,
             id24  TYPE c LENGTH 40,
             id25  TYPE c LENGTH 40,
             id26  TYPE c LENGTH 40,
             id27  TYPE c LENGTH 40,
             id28  TYPE c LENGTH 40,
             id30  TYPE c LENGTH 120,
           END OF lty_template.

    DATA et_template      TYPE TABLE OF lty_template.
    DATA et_template_temp TYPE TABLE OF lty_template.
    DATA es_template      TYPE lty_template.

    METHODS Excel_create.

ENDCLASS.


CLASS zcl_excel_create IMPLEMENTATION.
  METHOD excel_create.
    " XLSX document
    DATA lo_document   TYPE REF TO if_xco_xlsx_wa_document.
    DATA lo_sheet      TYPE REF TO if_xco_xlsx_wa_worksheet.
    DATA lv_sheet_name TYPE string.

    " Color Coding in the XLSX Sheet.
    TYPES to_color TYPE REF TO cl_xco_xlsx_color.

    DATA ls_bgcolor          TYPE to_color.
    DATA ls_fgcolor          TYPE to_color.

    DATA ls_horizontal_align TYPE REF TO cl_xco_xlsx_horzntal_alignment.

    lo_document = xco_cp_xlsx=>document->empty( )->write_access( ).
    " ----------------------------------------------------------------------
    " SHEET 1: Claim Data
    " ----------------------------------------------------------------------
    lv_sheet_name = |Field List|.
    lo_sheet = lo_document->get_workbook( )->worksheet->at_position( 1 ).
    lo_sheet->set_name( iv_name = lv_sheet_name ).

    " Fill et_result (first sheet content)
    CLEAR et_template.

    et_template_temp = VALUE #(
        " ------------------ SHEET TITLE AND NOTES ------------------
        ( id1 = 'Template Creation Field List'
          Id2 = ''
          id3 = ''
          id4 = '' )

        ( id1 = 'This sheet should be read-only.It provides a list of fields that are either required (marked with an asterisk) or optional for an import.'
          Id2 = ''
          id3 = ''
          id4 = '' )

        ( id1 = 'Field names are unique technical identifiers.They serve as a basis for data import.'
          Id2 = ''
          id3 = ''
          id4 = '' )

        " ------------------ HEADER COLUMNS ------------------
        ( id1 = 'Fields'
          Id2 = ''
          id3 = ''
          id4 = '' )

        ( id1 = 'Field Name (Technical)'
          Id2 = 'Field Description'
          id3 = ''
          id4 = 'Max. Length'
          id5 = 'Remarks' )

        " ------------------ HEADER FIELDS ------------------
        ( id1 = 'ID1'
          Id2 = 'ID1 Desc'
          id3 = 'Mandatory'
          id4 = '4'
          id5 = | User must enter a valid ID. Supported values are:{ cl_abap_char_utilities=>cr_lf }| &&
                      |• SAP{ cl_abap_char_utilities=>cr_lf }| &&
                      |• BTP | )

        ( id1 = 'ID2'
          Id2 = 'Date'
          id3 = 'Mandatory'
          id5 = |• Supports MM/DD/YYYY{ cl_abap_char_utilities=>cr_lf }| &&
                      |• Future date  is not allowed | )

        ( id1 = 'ID3'
          Id2 = 'ID3 Desc'
          id3 = 'Optional'
          id5 = |Enter the ID that belongs to the correct data.{ cl_abap_char_utilities=>cr_lf }| &&
                      |The ID will be created against the item.| )

        ( id1 = 'ID4'
          Id2 = 'ID4 Description'
          id3 = 'Mandatory'
          id4 = '4'
          id5 = |• Identifies the ID4.{ cl_abap_char_utilities=>cr_lf }| &&
                |• Determines the applicable values. | )

        ( id1 = 'ID5'
          Id2 = 'ID5 Description'
          id3 = 'Mandatory'
          id4 = '2'
          id5 = |• Identifies the ID5.{ cl_abap_char_utilities=>cr_lf }| &&
                |• Determines the applicable values. | )

        ( id1 = 'ID6'
          Id2 = 'ID6 Description'
          id3 = 'Mandatory'
          id4 = '2'
          id5 = |• Identifies the ID6.{ cl_abap_char_utilities=>cr_lf }| &&
                |• Determines the applicable values. | )

        ( id1 = 'ID7'
          Id2 = 'ID7 Description'
          id3 = 'Optional'
          id4 = '4'
          id5 = |• Identifies the ID7.{ cl_abap_char_utilities=>cr_lf }| &&
                |• Determines the applicable values. | )

        ( id1 = 'ID8'
          Id2 = 'ID8 Description'
          id3 = 'Optional'
          id4 = '4'
          id5 = |• Identifies the ID8.{ cl_abap_char_utilities=>cr_lf }| &&
                |• Determines the applicable values. | )

        ( id1 = 'ID9'
          Id2 = 'ID9 Description'
          id3 = 'Mandatory'
          id4 = '10'
          id5 = |• Identifies the ID9.{ cl_abap_char_utilities=>cr_lf }| &&
                |• Determines the applicable values.{ cl_abap_char_utilities=>cr_lf }| &&
                |• Must be based on ID3. | )
        " ------------------ PARTNER FIELDS ------------------
        ( )
        ( id1 = 'Secondary Fields'
          Id2 = ''
          id3 = ''
          id4 = '' )

        ( id1 = 'Field Name (Technical)'
          Id2 = 'Field Description'
          id3 = ''
          id4 = 'Max. Length'
          id5 = 'Remarks' )

        ( id1 = 'ID10'
          Id2 = 'ID DESC'
          id3 = 'Mandatory'
          id4 = '10'
          id5 = |• Specifies the ID1.{ cl_abap_char_utilities=>cr_lf }| &&
                      |• Requires a numeric key.{ cl_abap_char_utilities=>cr_lf }| &&
                      |• Must exist in master data.| )

        ( id1 = 'ID11'
          Id2 = 'ID11 Desc'
          id3 = 'Mandatory'
          id4 = '10'
          id5 = |• Specifies the ID11 or entity{ cl_abap_char_utilities=>cr_lf }| &&
                      |• Requires a numeric key.{ cl_abap_char_utilities=>cr_lf }| &&
                      |• Must exist in organizational master data.| )

        ( id1 = 'ID12'
          Id2 = 'ID12 Desc'
          id3 = 'Optional'
          id4 = '10'
          id5 = |• Indicates ID12 created.{ cl_abap_char_utilities=>cr_lf }| &&
                      |• Requires a numeric key.{ cl_abap_char_utilities=>cr_lf }| &&
                      |• Must exist in data.| )

        ( id1 = 'ID13'
          Id2 = 'ID13 Desc'
          id3 = 'Optional'
          id4 = '10'
          id5 = |• Identifies the ID13 values.{ cl_abap_char_utilities=>cr_lf }| &&
                      |• Requires a numeric key.{ cl_abap_char_utilities=>cr_lf }| &&
                      |• Must exist data.| )

        ( id1 = 'ID14'
          Id2 = 'ID14 Desc'
          id3 = 'Optional'
          id4 = '10'
          id5 = |• Specifies the transaction.{ cl_abap_char_utilities=>cr_lf }| &&
                      |• Requires a numeric key.{ cl_abap_char_utilities=>cr_lf }| &&
                      |• Must exist in master data.| )

        ( id1 = 'ID15'
          Id2 = 'ID15 Desc'
          id3 = 'Optional'
          id4 = '10'
          id5 = |• Identifies the ID15.{ cl_abap_char_utilities=>cr_lf }| &&
                      |• Requires a numeric key.{ cl_abap_char_utilities=>cr_lf }| &&
                      |• This is not a standard data.| )

        ( id1 = 'ID16'
          Id2 = 'ID16 Desc'
          id3 = 'Optional'
          id4 = '10'
          id5 = |• Identifies the ID16{ cl_abap_char_utilities=>cr_lf }| &&
                      |• Requires a numeric key.{ cl_abap_char_utilities=>cr_lf }| &&
                      |• This is not a standard data.| )

        " ------------------ ITEM FIELDS ------------------
        ( )
        ( id1 = 'Third Fields'
          Id2 = ''
          id3 = ''
          id4 = '' )

        ( id1 = 'Field Name (Technical)'
          Id2 = 'Field Description'
          id3 = ''
          id4 = 'Max. Length'
          id5 = 'Remarks' )

        ( id1 = 'ID17'
          Id2 = 'ID17 Desc'
          id3 = 'Optional'
          id4 = '40'
          id5 = |• Indicates the ID17.{ cl_abap_char_utilities=>cr_lf }| &&
                      |• Requires a numeric key.| )

        ( id1 = 'ID18'
          Id2 = 'ID18'
          id3 = 'Optional'
          id4 = '9'
          id5 = |• Indicates the ID18.{ cl_abap_char_utilities=>cr_lf }| &&
                      |• Requires a numeric key.{ cl_abap_char_utilities=>cr_lf }| &&
                      |• If not explicitly provided then derive.| )

        ( id1 = 'ID19'
          Id2 = 'ID19 Desc'
          id3 = 'Mandatory'
          id4 = '15'
          id5 = |• Specifies the monetary value.{ cl_abap_char_utilities=>cr_lf }| &&
                      |• Must be consistent.| )

        ( id1 = 'ID20'
          Id2 = 'ID20 Desc'
          id3 = 'Mandatory'
          id4 = '3'
          id5 = |• Indicates the currency.{ cl_abap_char_utilities=>cr_lf }| &&
                      |• Must follow ISO 4217 standard (e.g., USD, EUR, INR).| )

        ( id1 = 'ID21'
          Id2 = 'ID21 Desc'
          id3 = 'Optional'
          id4 = '13'
          id5 = |• Specifies the quantity.{ cl_abap_char_utilities=>cr_lf }| &&
                      |• Must be a positive numeric value.| )

        ( id1 = 'ID22'
          Id2 = 'Unit'
          id3 = 'Optional'
          id4 = '3'
          id5 = |• Indicates the unit of measure in which the quantity is expressed (e.g., EA = Each, KG = Kilogram, L = Liter). { cl_abap_char_utilities=>cr_lf }| &&
                      |• Must align with the master data.| )

        ( id1 = 'ID23'
          Id2 = 'ID23 Desc'
          id3 = 'Optional'
          id4 = ''
          id5 = '• Supports MM/DD/YYYY' )

        ( id1 = 'ID24'
          Id2 = 'ID24 Desc'
          id3 = 'Optional'
          id4 = ''
          id5 = |• Supports MM/DD/YYYY{ cl_abap_char_utilities=>cr_lf } • Date must be in correct format| )

        ( id1 = 'ID25'
          Id2 = 'ID25 Desc'
          id3 = 'Optional'
          id4 = '4'
          id5 = |Defines the overall SAP BTP integration and configuration context, including the relevant application, service, environment, and process classification.{ cl_abap_char_utilities=>cr_lf }{ cl_abap_char_utilities=>cr_lf }| &&
           |For SAP BTP:{ cl_abap_char_utilities=>cr_lf }| &&
           |Supported Values are:{ cl_abap_char_utilities=>cr_lf }| &&
           |• SAP BTP – Business Technology Platform{ cl_abap_char_utilities=>cr_lf }| &&
           |• SAP BTP Cockpit – Cloud-based administration and monitoring{ cl_abap_char_utilities=>cr_lf }| &&
           |• SAP BTP Integration Suite – Integration and API management capabilities{ cl_abap_char_utilities=>cr_lf }| &&
           |• SAP BTP Cloud Foundry – Application runtime environment{ cl_abap_char_utilities=>cr_lf }| &&
           |• SAP BTP Kyma – Kubernetes-based application runtime{ cl_abap_char_utilities=>cr_lf }{ cl_abap_char_utilities=>cr_lf }| &&
           |For SAP BTP Services:{ cl_abap_char_utilities=>cr_lf }| &&
           |Supported Values are:{ cl_abap_char_utilities=>cr_lf }| &&
           |• Integration Suite{ cl_abap_char_utilities=>cr_lf }| &&
           |• API Management{ cl_abap_char_utilities=>cr_lf }| &&
           |• SAP HANA Cloud{ cl_abap_char_utilities=>cr_lf }| &&
           |• SAP Build Apps{ cl_abap_char_utilities=>cr_lf }| &&
           |• SAP Event Mesh{ cl_abap_char_utilities=>cr_lf }{ cl_abap_char_utilities=>cr_lf }| &&
           |Note: Only the relevant SAP BTP components and services applicable to the selected integration or business process will be accepted by the system.{ cl_abap_char_utilities=>cr_lf }| )
        ( id1 = 'ID26'
          Id2 = 'ID26 Desc'
          id3 = 'Optional'
          id4 = '4'
          id5 = |• Indicates the XCO library feature related to the development object.{ cl_abap_char_utilities=>cr_lf }| &&
                |• User must select the corresponding XCO API or library component as required.{ cl_abap_char_utilities=>cr_lf }| &&
                |• Based on the selected object type, use the appropriate XCO library interface.{ cl_abap_char_utilities=>cr_lf }| &&
                |• The required XCO library details can be found in:{ cl_abap_char_utilities=>cr_lf }| &&
                |SAP Help Portal \{ ABAP Development > XCO Library > API Documentation\} | ) ).

    " Write to sheet 1
    DATA(lo_pattern_field) = xco_cp_xlsx_selection=>pattern_builder->simple_from_to(
                        )->get_pattern( ).

    lo_sheet->select( lo_pattern_field
                    )->row_stream(
                    )->operation->write_from( REF #( et_template_temp )
                    )->execute( ).

    " ------------------  File Details  ------------------

    " Row 1
    ls_bgcolor = xco_cp_xlsx=>color->standard->black.
    ls_fgcolor =  xco_cp_xlsx=>color->standard->white.
    DATA(lo_cursor) = lo_sheet->cursor( io_column = xco_cp_xlsx=>coordinate->for_alphabetic_value( 'B' )
                                        io_row    = xco_cp_xlsx=>coordinate->for_numeric_value( 1 ) ).

    DO 5 TIMES.
      DATA(lo_cell) = lo_cursor->get_cell( ).
      lo_cell->apply_styles(
          VALUE #( ( xco_cp_xlsx=>style->fill( )->set_background_color( ls_bgcolor ) )
                   ( xco_cp_xlsx=>style->font( )->set_size( 25 )->set_color( io_color = ls_fgcolor ) ) ) ).
      lo_cursor->move_right( ).
    ENDDO.

    " add the column width
    lo_sheet->column( xco_cp_xlsx=>coordinate->for_alphabetic_value( 'B' ) )->set_width( iv_width = 120 ).
    lo_sheet->column( xco_cp_xlsx=>coordinate->for_numeric_value( 2 )
                       )->set_width( iv_width = 100 ).

    " Merge the  column and row 1
    DATA(lo_patternA) = xco_cp_xlsx_selection=>pattern_builder->simple_from_to(
                        )->from_column( xco_cp_xlsx=>coordinate->for_alphabetic_value( 'B' )
                        )->from_row( xco_cp_xlsx=>coordinate->for_numeric_value( 1 )
                        )->to_column( xco_cp_xlsx=>coordinate->for_alphabetic_value( 'F' )
                        )->to_row( xco_cp_xlsx=>coordinate->for_numeric_value( 1 )
                        )->get_pattern( ).

    lo_sheet->merge_cells( lo_patternA ).

    " Row 2
    CLEAR: lo_cursor,
           lo_cell,
           lo_patternA,
           ls_bgcolor,
           ls_fgcolor.

    ls_bgcolor = xco_cp_xlsx=>color->standard->black.
    ls_fgcolor = xco_cp_xlsx=>color->standard->white.
    lo_cursor = lo_sheet->cursor( io_column = xco_cp_xlsx=>coordinate->for_alphabetic_value( 'B' )
                                  io_row    = xco_cp_xlsx=>coordinate->for_numeric_value( 2 ) ).

    DO 5 TIMES.
      lo_cell = lo_cursor->get_cell( ).
      lo_cell->apply_styles(
          VALUE #( ( xco_cp_xlsx=>style->fill( )->set_background_color( ls_bgcolor ) )
                   ( xco_cp_xlsx=>style->font( )->set_size( 12 )->set_color( io_color = ls_fgcolor ) ) ) ).
      lo_cursor->move_right( ).
    ENDDO.

    " Merge the  column and row 2
    lo_patternA = xco_cp_xlsx_selection=>pattern_builder->simple_from_to(
                    )->from_column( xco_cp_xlsx=>coordinate->for_alphabetic_value( 'B' )
                    )->from_row( xco_cp_xlsx=>coordinate->for_numeric_value( 2 )
                    )->to_column( xco_cp_xlsx=>coordinate->for_alphabetic_value( 'F' )
                    )->to_row( xco_cp_xlsx=>coordinate->for_numeric_value( 2 )
                    )->get_pattern( ).

    lo_sheet->merge_cells( lo_patternA ).

    " Row 3
    CLEAR: lo_cursor,
           lo_cell,
           lo_patternA,
           ls_bgcolor,
           ls_fgcolor.

    ls_bgcolor = xco_cp_xlsx=>color->standard->black.
    ls_fgcolor = xco_cp_xlsx=>color->standard->white.
    lo_cursor = lo_sheet->cursor( io_column = xco_cp_xlsx=>coordinate->for_alphabetic_value( 'B' )
                                  io_row    = xco_cp_xlsx=>coordinate->for_numeric_value( 3 ) ).

    DO 5 TIMES.
      lo_cell = lo_cursor->get_cell( ).
      lo_cell->apply_styles(
          VALUE #( ( xco_cp_xlsx=>style->fill( )->set_background_color( ls_bgcolor ) )
                   ( xco_cp_xlsx=>style->font( )->set_size( 12 )->set_color( io_color = ls_fgcolor ) ) ) ).
      lo_cursor->move_right( ).
    ENDDO.

    " Merge the  column and row 2
    lo_patternA = xco_cp_xlsx_selection=>pattern_builder->simple_from_to(
                    )->from_column( xco_cp_xlsx=>coordinate->for_alphabetic_value( 'B' )
                    )->from_row( xco_cp_xlsx=>coordinate->for_numeric_value( 3 )
                    )->to_column( xco_cp_xlsx=>coordinate->for_alphabetic_value( 'F' )
                    )->to_row( xco_cp_xlsx=>coordinate->for_numeric_value( 3 )
                    )->get_pattern( ).

    lo_sheet->merge_cells( lo_patternA ).

    " ------------------ HEADER  ------------------
    " Row 4
    ls_horizontal_align = xco_cp_xlsx=>horizontal_alignment->center.
    CLEAR: lo_cursor,
           lo_cell,
           lo_patternA,
           ls_bgcolor,
           ls_fgcolor.

    ls_bgcolor = xco_cp_xlsx=>color->standard->light_blue.
    ls_fgcolor = xco_cp_xlsx=>color->standard->white.
    lo_cursor = lo_sheet->cursor( io_column = xco_cp_xlsx=>coordinate->for_alphabetic_value( 'B' )
                                  io_row    = xco_cp_xlsx=>coordinate->for_numeric_value( 4 ) ).

    DO 5 TIMES.
      lo_cell = lo_cursor->get_cell( ).
      lo_cell->apply_styles(
          VALUE #( ( xco_cp_xlsx=>style->fill( )->set_background_color( ls_bgcolor ) )
                   ( xco_cp_xlsx=>style->font( )->set_size( 12 )->set_color( io_color = ls_fgcolor )->set_bold( ) )
                   ( xco_cp_xlsx=>style->alignment( )->set_horizontal_alignment(
                         io_horizontal_alignment = ls_horizontal_align ) ) ) ).
      lo_cursor->move_right( ).
    ENDDO.

    " Merge the  column and row 2
    lo_patternA = xco_cp_xlsx_selection=>pattern_builder->simple_from_to(
                    )->from_column( xco_cp_xlsx=>coordinate->for_alphabetic_value( 'B' )
                    )->from_row( xco_cp_xlsx=>coordinate->for_numeric_value( 4 )
                    )->to_column( xco_cp_xlsx=>coordinate->for_alphabetic_value( 'F' )
                    )->to_row( xco_cp_xlsx=>coordinate->for_numeric_value( 4 )
                    )->get_pattern( ).

    lo_sheet->merge_cells( lo_patternA ).

    " Row 7
    CLEAR: lo_cursor,
           lo_cell,
           ls_bgcolor,
           ls_fgcolor.

    lo_sheet->column( xco_cp_xlsx=>coordinate->for_alphabetic_value( 'B' ) )->set_width( iv_width = 25 ).
    lo_sheet->column( xco_cp_xlsx=>coordinate->for_alphabetic_value( 'C' ) )->set_width( iv_width = 25 ).
    lo_sheet->column( xco_cp_xlsx=>coordinate->for_alphabetic_value( 'D' ) )->set_width( iv_width = 20 ).
    lo_sheet->column( xco_cp_xlsx=>coordinate->for_alphabetic_value( 'E' ) )->set_width( iv_width = 15 ).
    lo_sheet->column( xco_cp_xlsx=>coordinate->for_alphabetic_value( 'F' ) )->set_width( iv_width = 65 ).

    ls_bgcolor = xco_cp_xlsx=>color->standard->purple.
    ls_fgcolor = xco_cp_xlsx=>color->standard->white.
    lo_cursor = lo_sheet->cursor( io_column = xco_cp_xlsx=>coordinate->for_alphabetic_value( 'B' )
                                  io_row    = xco_cp_xlsx=>coordinate->for_numeric_value( 5 ) ).

    DO 5 TIMES.
      lo_cell = lo_cursor->get_cell( ).
      lo_cell->apply_styles(
          VALUE #( ( xco_cp_xlsx=>style->fill( )->set_background_color( ls_bgcolor ) )
                   ( xco_cp_xlsx=>style->font( )->set_size( 12 )->set_color( io_color = ls_fgcolor )->set_bold( ) )
                   ( xco_cp_xlsx=>style->alignment( )->set_horizontal_alignment(
                         io_horizontal_alignment = ls_horizontal_align ) ) ) ).

      lo_cursor->move_right( ).
    ENDDO.

    " Row 6 to 14   ======================= Header Fields ==============================================
    CLEAR: lo_cursor,
           lo_cell,
           ls_bgcolor,
           ls_fgcolor.

    DATA lv_count TYPE i VALUE 6.

    WHILE lv_count < 15.
      lo_cursor = lo_sheet->cursor( io_column = xco_cp_xlsx=>coordinate->for_alphabetic_value( 'B' )
                                    io_row    = xco_cp_xlsx=>coordinate->for_numeric_value( lv_count ) ).

      DO 5 TIMES.
        lo_cell = lo_cursor->get_cell( ).
        lo_cell->apply_styles(
            VALUE #(
                ( xco_cp_xlsx=>style->fill( ) )
                ( xco_cp_xlsx=>style->font( )->set_size( 10 ) )
                ( xco_cp_xlsx=>style->alignment( )->set_horizontal_alignment( xco_cp_xlsx=>horizontal_alignment->right
                 )->set_indent( 10 ) )
                ( xco_cp_xlsx=>style->alignment( )->set_vertical_alignment( xco_cp_xlsx=>vertical_alignment->center
                 )->set_wrap_text( ) ) ) ).

        lo_cursor->move_right( ).

      ENDDO.
      lv_count += 1.
    ENDWHILE.

*    " ------------------ PARTNER FIELDS ------------------

    " Row 16
    ls_horizontal_align = xco_cp_xlsx=>horizontal_alignment->center.
    CLEAR: lo_cursor,
           lo_cell,
           lo_patternA,
           ls_bgcolor,
           ls_fgcolor.

    ls_bgcolor = xco_cp_xlsx=>color->standard->blue.
    ls_fgcolor = xco_cp_xlsx=>color->standard->white.
    lo_cursor = lo_sheet->cursor( io_column = xco_cp_xlsx=>coordinate->for_alphabetic_value( 'B' )
                                  io_row    = xco_cp_xlsx=>coordinate->for_numeric_value( 16 ) ).

    DO 5 TIMES.
      lo_cell = lo_cursor->get_cell( ).
      lo_cell->apply_styles(
          VALUE #( ( xco_cp_xlsx=>style->fill( )->set_background_color( ls_bgcolor ) )
                   ( xco_cp_xlsx=>style->font( )->set_size( 12 )->set_color( io_color = ls_fgcolor )->set_bold( ) )
                   ( xco_cp_xlsx=>style->alignment( )->set_horizontal_alignment(
                         io_horizontal_alignment = ls_horizontal_align ) ) ) ).
      lo_cursor->move_right( ).
    ENDDO.

    " Merge the  column and row 2
    lo_patternA = xco_cp_xlsx_selection=>pattern_builder->simple_from_to(
                    )->from_column( xco_cp_xlsx=>coordinate->for_alphabetic_value( 'B' )
                    )->from_row( xco_cp_xlsx=>coordinate->for_numeric_value( 16 )
                    )->to_column( xco_cp_xlsx=>coordinate->for_alphabetic_value( 'F' )
                    )->to_row( xco_cp_xlsx=>coordinate->for_numeric_value( 16 )
                    )->get_pattern( ).

    lo_sheet->merge_cells( lo_patternA ).

    " row 17
    CLEAR: lo_cursor,
           lo_cell,
           ls_bgcolor,
           ls_fgcolor.

    lo_sheet->column( xco_cp_xlsx=>coordinate->for_alphabetic_value( 'B' ) )->set_width( iv_width = 25 ).
    lo_sheet->column( xco_cp_xlsx=>coordinate->for_alphabetic_value( 'C' ) )->set_width( iv_width = 25 ).
    lo_sheet->column( xco_cp_xlsx=>coordinate->for_alphabetic_value( 'D' ) )->set_width( iv_width = 20 ).
    lo_sheet->column( xco_cp_xlsx=>coordinate->for_alphabetic_value( 'E' ) )->set_width( iv_width = 15 ).
    lo_sheet->column( xco_cp_xlsx=>coordinate->for_alphabetic_value( 'F' ) )->set_width( iv_width = 125 ).

    ls_bgcolor = xco_cp_xlsx=>color->standard->purple.
    ls_fgcolor = xco_cp_xlsx=>color->standard->white.
    lo_cursor = lo_sheet->cursor( io_column = xco_cp_xlsx=>coordinate->for_alphabetic_value( 'B' )
                                  io_row    = xco_cp_xlsx=>coordinate->for_numeric_value( 17 ) ).

    DO 5 TIMES.
      lo_cell = lo_cursor->get_cell( ).
      lo_cell->apply_styles(
          VALUE #( ( xco_cp_xlsx=>style->fill( )->set_background_color( ls_bgcolor ) )
                   ( xco_cp_xlsx=>style->font( )->set_size( 12 )->set_color( io_color = ls_fgcolor )->set_bold( ) )
                   ( xco_cp_xlsx=>style->alignment( )->set_horizontal_alignment(
                         io_horizontal_alignment = ls_horizontal_align ) ) ) ).

      lo_cursor->move_right( ).
    ENDDO.

    " row 18 to 23
    CLEAR: lo_cursor,
           lo_cell,
           ls_bgcolor,
           ls_fgcolor.

    lv_count = 17.

    WHILE lv_count < 25.
      lo_cursor = lo_sheet->cursor( io_column = xco_cp_xlsx=>coordinate->for_alphabetic_value( 'B' )
                                    io_row    = xco_cp_xlsx=>coordinate->for_numeric_value( lv_count ) ).

      DO 5 TIMES.
        lo_cell = lo_cursor->get_cell( ).
        lo_cell->apply_styles(
            VALUE #(
                ( xco_cp_xlsx=>style->fill( ) )
                ( xco_cp_xlsx=>style->font( )->set_size( 10 ) )
                ( xco_cp_xlsx=>style->alignment( )->set_horizontal_alignment( xco_cp_xlsx=>horizontal_alignment->right
                 )->set_indent( 10 ) )
                ( xco_cp_xlsx=>style->alignment( )->set_vertical_alignment( xco_cp_xlsx=>vertical_alignment->center
                 )->set_wrap_text( ) ) ) ).

        lo_cursor->move_right( ).

      ENDDO.
      lv_count += 1.
    ENDWHILE.

    " ------------------ ITEM FIELDS ------------------
    " Row 24
    ls_horizontal_align = xco_cp_xlsx=>horizontal_alignment->center.
    CLEAR: lo_cursor,
           lo_cell,
           lo_patternA,
           ls_bgcolor,
           ls_fgcolor.

    ls_bgcolor = xco_cp_xlsx=>color->standard->dark_blue.
    ls_fgcolor = xco_cp_xlsx=>color->standard->white.
    lo_cursor = lo_sheet->cursor( io_column = xco_cp_xlsx=>coordinate->for_alphabetic_value( 'B' )
                                  io_row    = xco_cp_xlsx=>coordinate->for_numeric_value( 26 ) ).

    DO 5 TIMES.
      lo_cell = lo_cursor->get_cell( ).
      lo_cell->apply_styles(
          VALUE #( ( xco_cp_xlsx=>style->fill( )->set_background_color( ls_bgcolor ) )
                   ( xco_cp_xlsx=>style->font( )->set_size( 12 )->set_color( io_color = ls_fgcolor )->set_bold( ) )
                   ( xco_cp_xlsx=>style->alignment( )->set_horizontal_alignment(
                         io_horizontal_alignment = ls_horizontal_align ) ) ) ).
      lo_cursor->move_right( ).
    ENDDO.

    " Merge the  column and row 2
    lo_patternA = xco_cp_xlsx_selection=>pattern_builder->simple_from_to(
                    )->from_column( xco_cp_xlsx=>coordinate->for_alphabetic_value( 'B' )
                    )->from_row( xco_cp_xlsx=>coordinate->for_numeric_value( 26 )
                    )->to_column( xco_cp_xlsx=>coordinate->for_alphabetic_value( 'F' )
                    )->to_row( xco_cp_xlsx=>coordinate->for_numeric_value( 26 )
                    )->get_pattern( ).

    lo_sheet->merge_cells( lo_patternA ).

    " row 27
    CLEAR: lo_cursor,
           lo_cell,
           ls_bgcolor,
           ls_fgcolor.

    lo_sheet->column( xco_cp_xlsx=>coordinate->for_alphabetic_value( 'B' ) )->set_width( iv_width = 25 ).
    lo_sheet->column( xco_cp_xlsx=>coordinate->for_alphabetic_value( 'C' ) )->set_width( iv_width = 25 ).
    lo_sheet->column( xco_cp_xlsx=>coordinate->for_alphabetic_value( 'D' ) )->set_width( iv_width = 20 ).
    lo_sheet->column( xco_cp_xlsx=>coordinate->for_alphabetic_value( 'E' ) )->set_width( iv_width = 15 ).
    lo_sheet->column( xco_cp_xlsx=>coordinate->for_alphabetic_value( 'F' ) )->set_width( iv_width = 125 ).

    ls_bgcolor = xco_cp_xlsx=>color->standard->purple.
    ls_fgcolor = xco_cp_xlsx=>color->standard->white.
    lo_cursor = lo_sheet->cursor( io_column = xco_cp_xlsx=>coordinate->for_alphabetic_value( 'B' )
                                  io_row    = xco_cp_xlsx=>coordinate->for_numeric_value( 27 ) ).

    DO 5 TIMES.
      lo_cell = lo_cursor->get_cell( ).
      lo_cell->apply_styles(
          VALUE #( ( xco_cp_xlsx=>style->fill( )->set_background_color( ls_bgcolor ) )
                   ( xco_cp_xlsx=>style->font( )->set_size( 12 )->set_color( io_color = ls_fgcolor )->set_bold( ) )
                   ( xco_cp_xlsx=>style->alignment( )->set_horizontal_alignment(
                         io_horizontal_alignment = ls_horizontal_align ) ) ) ).

      lo_cursor->move_right( ).
    ENDDO.

    " row 27 to 37
    CLEAR: lo_cursor,
           lo_cell,
           ls_bgcolor,
           ls_fgcolor.

    lv_count = 28.

    WHILE lv_count < 38.
      lo_cursor = lo_sheet->cursor( io_column = xco_cp_xlsx=>coordinate->for_alphabetic_value( 'B' )
                                    io_row    = xco_cp_xlsx=>coordinate->for_numeric_value( lv_count ) ).

      DO 5 TIMES.
        lo_cell = lo_cursor->get_cell( ).
        lo_cell->apply_styles(
            VALUE #(
                ( xco_cp_xlsx=>style->fill( ) )
                ( xco_cp_xlsx=>style->font( )->set_size( 10 ) )
                ( xco_cp_xlsx=>style->alignment( )->set_horizontal_alignment( xco_cp_xlsx=>horizontal_alignment->right
                 )->set_indent( 10 ) )
                ( xco_cp_xlsx=>style->alignment( )->set_vertical_alignment( xco_cp_xlsx=>vertical_alignment->center
                 )->set_wrap_text( ) ) ) ).

        lo_cursor->move_right( ).

      ENDDO.
      lv_count += 1.
    ENDWHILE.

    " Restrict customer to change the things in Sheet 1.
    lo_sheet->protect( ).

    " ---------------------------------------------------------------------
    " SHEET 2: Claim Template
    " ----------------------------------------------------------------------
    CLEAR: et_template,
           lv_sheet_name,
           lo_sheet.
    lv_sheet_name = |2ND Worksheet|.
    lo_sheet = lo_document->get_workbook( )->add_new_sheet( ). " create 2nd sheet
    lo_sheet->set_name( iv_name = lv_sheet_name ).

    et_template = VALUE #(
        ( space = 'ID'
          id1   = 'Fill in Request Data' )

        ( space = 'File Number'
          id1   = |Double Click this cell for more description{ cl_abap_char_utilities=>cr_lf }{ cl_abap_char_utilities=>cr_lf }| &&
                     |Notes About Filling in Claim Request Data{ cl_abap_char_utilities=>cr_lf }| &&
                     |-Enter keys for non-description fields. For example, if Promotion Trade Spend Type has been configured as “ZT-0013 Scan-back Allowance” in your configuration environment,{ cl_abap_char_utilities=>cr_lf }|
&&
                      |enter ZT-0013 in the Promotion Trade Spend Type field.{ cl_abap_char_utilities=>cr_lf }| &&
                      |-To query possible key values of these fields, refer to the Field List sheet or SAP master data configuration{ cl_abap_char_utilities=>cr_lf }| &&
                      |-A maximum of 1,000 records is recommended per import.{ cl_abap_char_utilities=>cr_lf }| )

        ( Id2   = 'Fields'
          id11  = 'Second Fields'
          id18  = 'Third Fields'
          id28  = 'Not For Input' )

        ( Id2   = 'ID1'
          id3   = 'ID2'
          id4   = 'ID3'
          id5   = 'ID4'
          id6   = 'ID5'
          id7   = 'ID6'
          id8   = 'ID7'
          id9   = 'ID8'
          id10  = 'ID9'
          id11  = 'ID10'
          id12  = 'ID10'
          id13  = 'ID11'
          id14  = 'ID12'
          id15  = 'ID13'
          id16  = 'ID14'
          id17  = 'ID15'
          id18  = 'ID16'
          id19  = 'ID17'
          id20  = 'ID18'
          id21  = 'ID19'
          id22  = 'ID20'
          id23  = 'ID21'
          id24  = 'ID22'
          id25  = 'ID23'
          id26  = 'ID24'
          id27  = 'ID25'
          id28  = 'ID26'
          id30  = 'ID27' )

        ( id1   = 'ID1'
          id2   = '*ID2'
          id3   = '*ID3'
          id4   = 'ID4'
          id5   = '*ID5'
          id6   = '*ID6'
          id7   = '*ID7'
          id8   = 'ID8'
          id9   = 'ID9'
          id10  = '*ID10'
          id11  = '*ID11'
          id12  = '*ID12'
          id13  = 'ID13'
          id14  = 'ID14'
          id15  = 'ID15'
          id16  = 'ID16'
          id17  = 'ID17'
          id18  = 'ID18'
          id19  = 'ID19'
          id20  = '*ID20'
          id21  = '*ID21'
          id22  = 'ID22'
          id23  = 'ID23'
          id24  = 'ID24'
          id25  = 'ID25'
          id26  = 'ID26'
          id27  = 'ID27'
          id28  = 'ID28'
          id30  = 'ID30' ) ).

    DATA(lo_pattern) = xco_cp_xlsx_selection=>pattern_builder->simple_from_to( )->get_pattern( ).

    lo_sheet->select( lo_pattern
      )->row_stream(
      )->operation->write_from( REF #( et_template )
      )->execute( ).

    " Styling for the row and column of second excel sheet
    " Column A --> Row1 ===> A1
    ls_horizontal_align = xco_cp_xlsx=>horizontal_alignment->center.
    CLEAR: lo_cursor,
           lo_cell,
           lo_patternA,
           ls_bgcolor,
           ls_fgcolor.

    lo_sheet->column( xco_cp_xlsx=>coordinate->for_alphabetic_value( 'A' ) )->set_width( iv_width = 12 ).
    ls_bgcolor = xco_cp_xlsx=>color->standard->purple.
    ls_fgcolor = xco_cp_xlsx=>color->standard->white.
    lo_cursor = lo_sheet->cursor( io_column = xco_cp_xlsx=>coordinate->for_alphabetic_value( 'A' )
                                  io_row    = xco_cp_xlsx=>coordinate->for_numeric_value( 1 ) ).
    lo_cell = lo_cursor->get_cell( ).
    lo_cell->apply_styles(
        VALUE #(
            ( xco_cp_xlsx=>style->fill( )->set_background_color( ls_bgcolor ) )
            ( xco_cp_xlsx=>style->font( )->set_bold( )->set_size( 11 )->set_color( io_color = ls_fgcolor ) )
            ( xco_cp_xlsx=>style->alignment( )->set_horizontal_alignment( xco_cp_xlsx=>horizontal_alignment->left ) )
            ( xco_cp_xlsx=>style->alignment( )->set_vertical_alignment( xco_cp_xlsx=>vertical_alignment->center ) ) ) ).

    " Column B ==> B1
    " Merge the Columns Till --> B1 to AF
    ls_horizontal_align = xco_cp_xlsx=>horizontal_alignment->center.
    CLEAR: lo_cursor,
           lo_cell,
           lo_patternA,
           ls_bgcolor,
           ls_fgcolor.

*    lo_sheet->column( xco_cp_xlsx=>coordinate->for_alphabetic_value( 'B' ) )->set_width( iv_width = 12 ).
    ls_bgcolor = xco_cp_xlsx=>color->standard->Black.
    ls_fgcolor = xco_cp_xlsx=>color->standard->white.
    lo_cursor = lo_sheet->cursor( io_column = xco_cp_xlsx=>coordinate->for_alphabetic_value( 'B' )
                                  io_row    = xco_cp_xlsx=>coordinate->for_numeric_value( 1 ) ).
    lo_cell = lo_cursor->get_cell( ).
    lo_cell->apply_styles(
        VALUE #(
            ( xco_cp_xlsx=>style->fill( )->set_background_color( ls_bgcolor ) )
            ( xco_cp_xlsx=>style->font( )->set_bold( )->set_size( 25 )->set_color( io_color = ls_fgcolor ) )
            ( xco_cp_xlsx=>style->alignment( )->set_horizontal_alignment( xco_cp_xlsx=>horizontal_alignment->left ) )
            ( xco_cp_xlsx=>style->alignment( )->set_vertical_alignment( xco_cp_xlsx=>vertical_alignment->center ) ) ) ).

    lo_patternA = xco_cp_xlsx_selection=>pattern_builder->simple_from_to(
               )->from_column( xco_cp_xlsx=>coordinate->for_alphabetic_value( 'B' )
               )->from_row( xco_cp_xlsx=>coordinate->for_numeric_value( 1 )
               )->to_column( xco_cp_xlsx=>coordinate->for_alphabetic_value( 'AD' )
               )->to_row( xco_cp_xlsx=>coordinate->for_numeric_value( 1 )
               )->get_pattern( ).

    lo_sheet->merge_cells( lo_patternA ).

    " Column B ==> B2
    " Merge the Columns Till --> B2 to AF
    ls_horizontal_align = xco_cp_xlsx=>horizontal_alignment->center.
    CLEAR: lo_cursor,
           lo_cell,
           lo_patternA,
           ls_bgcolor,
           ls_fgcolor.

    ls_bgcolor = xco_cp_xlsx=>color->standard->Purple.
    ls_fgcolor = xco_cp_xlsx=>color->standard->white.
    lo_cursor = lo_sheet->cursor( io_column = xco_cp_xlsx=>coordinate->for_alphabetic_value( 'B' )
                                  io_row    = xco_cp_xlsx=>coordinate->for_numeric_value( 2 ) ).
    lo_cell = lo_cursor->get_cell( ).
    lo_cell->apply_styles(
        VALUE #(
            ( xco_cp_xlsx=>style->fill( )->set_background_color( ls_bgcolor ) )
            ( xco_cp_xlsx=>style->font( )->set_bold( )->set_size( 12 )->set_color( io_color = ls_fgcolor ) )
            ( xco_cp_xlsx=>style->alignment( )->set_vertical_alignment( xco_cp_xlsx=>vertical_alignment->center ) )
            ( xco_cp_xlsx=>style->alignment( )->set_horizontal_alignment( xco_cp_xlsx=>horizontal_alignment->left )->set_wrap_text( ) ) ) ).

    lo_patternA = xco_cp_xlsx_selection=>pattern_builder->simple_from_to(
               )->from_column( xco_cp_xlsx=>coordinate->for_alphabetic_value( 'B' )
               )->from_row( xco_cp_xlsx=>coordinate->for_numeric_value( 2 )
               )->to_column( xco_cp_xlsx=>coordinate->for_alphabetic_value( 'AD' )
               )->to_row( xco_cp_xlsx=>coordinate->for_numeric_value( 2 )
               )->get_pattern( ).

    lo_sheet->merge_cells( lo_patternA ).

    " Header Name --> From B3 to R3
    ls_horizontal_align = xco_cp_xlsx=>horizontal_alignment->center.
    CLEAR: lo_cursor,
           lo_cell,
           lo_patternA,
           ls_bgcolor,
           ls_fgcolor.

    ls_bgcolor = xco_cp_xlsx=>color->standard->light_blue.
    ls_fgcolor = xco_cp_xlsx=>color->standard->white.
    lo_cursor = lo_sheet->cursor( io_column = xco_cp_xlsx=>coordinate->for_alphabetic_value( 'C' )
                                  io_row    = xco_cp_xlsx=>coordinate->for_numeric_value( 3 ) ).
    lo_cell = lo_cursor->get_cell( ).
    lo_cell->apply_styles(
        VALUE #(
            ( xco_cp_xlsx=>style->fill( )->set_background_color( ls_bgcolor ) )
            ( xco_cp_xlsx=>style->font( )->set_bold( )->set_size( 12 )->set_color( io_color = ls_fgcolor ) )
            ( xco_cp_xlsx=>style->alignment( )->set_vertical_alignment( xco_cp_xlsx=>vertical_alignment->center ) )
            ( xco_cp_xlsx=>style->alignment( )->set_horizontal_alignment( xco_cp_xlsx=>horizontal_alignment->center ) ) ) ).

    lo_patternA = xco_cp_xlsx_selection=>pattern_builder->simple_from_to(
               )->from_column( xco_cp_xlsx=>coordinate->for_alphabetic_value( 'C' )
               )->from_row( xco_cp_xlsx=>coordinate->for_numeric_value( 3 )
               )->to_column( xco_cp_xlsx=>coordinate->for_alphabetic_value( 'K' )
               )->to_row( xco_cp_xlsx=>coordinate->for_numeric_value( 3 )
               )->get_pattern( ).

    lo_sheet->merge_cells( lo_patternA ).

    " Partner Name
    ls_horizontal_align = xco_cp_xlsx=>horizontal_alignment->center.
    CLEAR: lo_cursor,
           lo_cell,
           lo_patternA,
           ls_bgcolor,
           ls_fgcolor.

    ls_bgcolor = xco_cp_xlsx=>color->standard->blue.
    ls_fgcolor = xco_cp_xlsx=>color->standard->white.
    lo_cursor = lo_sheet->cursor( io_column = xco_cp_xlsx=>coordinate->for_alphabetic_value( 'L' )
                                  io_row    = xco_cp_xlsx=>coordinate->for_numeric_value( 3 ) ).
    lo_cell = lo_cursor->get_cell( ).
    lo_cell->apply_styles(
        VALUE #(
            ( xco_cp_xlsx=>style->fill( )->set_background_color( ls_bgcolor ) )
            ( xco_cp_xlsx=>style->font( )->set_bold( )->set_size( 12 )->set_color( io_color = ls_fgcolor ) )
            ( xco_cp_xlsx=>style->alignment( )->set_vertical_alignment( xco_cp_xlsx=>vertical_alignment->center ) )
            ( xco_cp_xlsx=>style->alignment( )->set_horizontal_alignment( xco_cp_xlsx=>horizontal_alignment->center ) ) ) ).

    lo_patternA = xco_cp_xlsx_selection=>pattern_builder->simple_from_to(
               )->from_column( xco_cp_xlsx=>coordinate->for_alphabetic_value( 'L' )
               )->from_row( xco_cp_xlsx=>coordinate->for_numeric_value( 3 )
               )->to_column( xco_cp_xlsx=>coordinate->for_alphabetic_value( 'R' )
               )->to_row( xco_cp_xlsx=>coordinate->for_numeric_value( 3 )
               )->get_pattern( ).

    lo_sheet->merge_cells( lo_patternA ).

    " Item Name
    ls_horizontal_align = xco_cp_xlsx=>horizontal_alignment->center.
    CLEAR: lo_cursor,
           lo_cell,
           lo_patternA,
           ls_bgcolor,
           ls_fgcolor.

    ls_bgcolor = xco_cp_xlsx=>color->standard->dark_blue.
    ls_fgcolor = xco_cp_xlsx=>color->standard->white.
    lo_cursor = lo_sheet->cursor( io_column = xco_cp_xlsx=>coordinate->for_alphabetic_value( 'S' )
                                  io_row    = xco_cp_xlsx=>coordinate->for_numeric_value( 3 ) ).
    lo_cell = lo_cursor->get_cell( ).
    lo_cell->apply_styles(
        VALUE #(
            ( xco_cp_xlsx=>style->fill( )->set_background_color( ls_bgcolor ) )
            ( xco_cp_xlsx=>style->font( )->set_bold( )->set_size( 12 )->set_color( io_color = ls_fgcolor ) )
            ( xco_cp_xlsx=>style->alignment( )->set_vertical_alignment( xco_cp_xlsx=>vertical_alignment->center ) )
            ( xco_cp_xlsx=>style->alignment( )->set_horizontal_alignment( xco_cp_xlsx=>horizontal_alignment->center ) ) ) ).

    lo_patternA = xco_cp_xlsx_selection=>pattern_builder->simple_from_to(
               )->from_column( xco_cp_xlsx=>coordinate->for_alphabetic_value( 'S' )
               )->from_row( xco_cp_xlsx=>coordinate->for_numeric_value( 3 )
               )->to_column( xco_cp_xlsx=>coordinate->for_alphabetic_value( 'AB' )
               )->to_row( xco_cp_xlsx=>coordinate->for_numeric_value( 3 )
               )->get_pattern( ).

    lo_sheet->merge_cells( lo_patternA ).

    " Not Input Name
    ls_horizontal_align = xco_cp_xlsx=>horizontal_alignment->center.
    CLEAR: lo_cursor,
           lo_cell,
           lo_patternA,
           ls_bgcolor,
           ls_fgcolor.

    ls_bgcolor = xco_cp_xlsx=>color->standard->red.
    ls_fgcolor = xco_cp_xlsx=>color->standard->white.
    lo_cursor = lo_sheet->cursor( io_column = xco_cp_xlsx=>coordinate->for_alphabetic_value( 'AC' )
                                  io_row    = xco_cp_xlsx=>coordinate->for_numeric_value( 3 ) ).
    lo_cell = lo_cursor->get_cell( ).
    lo_cell->apply_styles(
        VALUE #(
            ( xco_cp_xlsx=>style->fill( )->set_background_color( ls_bgcolor ) )
            ( xco_cp_xlsx=>style->font( )->set_bold( )->set_size( 12 )->set_color( io_color = ls_fgcolor ) )
            ( xco_cp_xlsx=>style->alignment( )->set_vertical_alignment( xco_cp_xlsx=>vertical_alignment->center ) )
            ( xco_cp_xlsx=>style->alignment( )->set_horizontal_alignment( xco_cp_xlsx=>horizontal_alignment->center ) ) ) ).

    lo_patternA = xco_cp_xlsx_selection=>pattern_builder->simple_from_to(
               )->from_column( xco_cp_xlsx=>coordinate->for_alphabetic_value( 'AC' )
               )->from_row( xco_cp_xlsx=>coordinate->for_numeric_value( 3 )
               )->to_column( xco_cp_xlsx=>coordinate->for_alphabetic_value( 'AD' )
               )->to_row( xco_cp_xlsx=>coordinate->for_numeric_value( 3 )
               )->get_pattern( ).

    lo_sheet->merge_cells( lo_patternA ).

    " For the data present in the row 5.

    CLEAR: lo_cursor,
           lo_cell,
           ls_bgcolor,
           ls_fgcolor.

    lo_sheet->column( xco_cp_xlsx=>coordinate->for_alphabetic_value( 'B' ) )->set_width( iv_width = 15 ).
    lo_sheet->column( xco_cp_xlsx=>coordinate->for_alphabetic_value( 'C' ) )->set_width( iv_width = 15 ).
    lo_sheet->column( xco_cp_xlsx=>coordinate->for_alphabetic_value( 'D' ) )->set_width( iv_width = 20 ).
    lo_sheet->column( xco_cp_xlsx=>coordinate->for_alphabetic_value( 'E' ) )->set_width( iv_width = 23 ).
    lo_sheet->column( xco_cp_xlsx=>coordinate->for_alphabetic_value( 'F' ) )->set_width( iv_width = 23 ).
    lo_sheet->column( xco_cp_xlsx=>coordinate->for_alphabetic_value( 'G' ) )->set_width( iv_width = 23 ).
    lo_sheet->column( xco_cp_xlsx=>coordinate->for_alphabetic_value( 'H' ) )->set_width( iv_width = 18 ).
    lo_sheet->column( xco_cp_xlsx=>coordinate->for_alphabetic_value( 'I' ) )->set_width( iv_width = 18 ).
    lo_sheet->column( xco_cp_xlsx=>coordinate->for_alphabetic_value( 'J' ) )->set_width( iv_width = 18 ).
    lo_sheet->column( xco_cp_xlsx=>coordinate->for_alphabetic_value( 'K' ) )->set_width( iv_width = 18 ).
    lo_sheet->column( xco_cp_xlsx=>coordinate->for_alphabetic_value( 'L' ) )->set_width( iv_width = 18 ).
    lo_sheet->column( xco_cp_xlsx=>coordinate->for_alphabetic_value( 'M' ) )->set_width( iv_width = 18 ).
    lo_sheet->column( xco_cp_xlsx=>coordinate->for_alphabetic_value( 'N' ) )->set_width( iv_width = 22 ).
    lo_sheet->column( xco_cp_xlsx=>coordinate->for_alphabetic_value( 'O' ) )->set_width( iv_width = 18 ).
    lo_sheet->column( xco_cp_xlsx=>coordinate->for_alphabetic_value( 'P' ) )->set_width( iv_width = 18 ).
    lo_sheet->column( xco_cp_xlsx=>coordinate->for_alphabetic_value( 'Q' ) )->set_width( iv_width = 27 ).
    lo_sheet->column( xco_cp_xlsx=>coordinate->for_alphabetic_value( 'R' ) )->set_width( iv_width = 20 ).
    lo_sheet->column( xco_cp_xlsx=>coordinate->for_alphabetic_value( 'S' ) )->set_width( iv_width = 20 ).
    lo_sheet->column( xco_cp_xlsx=>coordinate->for_alphabetic_value( 'T' ) )->set_width( iv_width = 18 ).
    lo_sheet->column( xco_cp_xlsx=>coordinate->for_alphabetic_value( 'U' ) )->set_width( iv_width = 15 ).
    lo_sheet->column( xco_cp_xlsx=>coordinate->for_alphabetic_value( 'V' ) )->set_width( iv_width = 15 ).
    lo_sheet->column( xco_cp_xlsx=>coordinate->for_alphabetic_value( 'W' ) )->set_width( iv_width = 15 ).
    lo_sheet->column( xco_cp_xlsx=>coordinate->for_alphabetic_value( 'X' ) )->set_width( iv_width = 15 ).
    lo_sheet->column( xco_cp_xlsx=>coordinate->for_alphabetic_value( 'Y' ) )->set_width( iv_width = 23 ).
    lo_sheet->column( xco_cp_xlsx=>coordinate->for_alphabetic_value( 'Z' ) )->set_width( iv_width = 23 ).
    lo_sheet->column( xco_cp_xlsx=>coordinate->for_alphabetic_value( 'AA' ) )->set_width( iv_width = 22 ).
    lo_sheet->column( xco_cp_xlsx=>coordinate->for_alphabetic_value( 'AB' ) )->set_width( iv_width = 22 ).
    lo_sheet->column( xco_cp_xlsx=>coordinate->for_alphabetic_value( 'AC' ) )->set_width( iv_width = 22 ).
    lo_sheet->column( xco_cp_xlsx=>coordinate->for_alphabetic_value( 'AD' ) )->set_width( iv_width = 30 ).

    DATA(ld_excel) = lo_document->get_file_content( ).
    " TODO: variable is assigned but only used in commented-out code (ABAP cleaner)
    DATA(lv_base64_encoding) = xco_cp=>xstring( ld_excel )->as_string( xco_cp_binary=>text_encoding->base64 )->value.

  ENDMETHOD.
ENDCLASS.