CLASS zcl_excel_upload DEFINITION
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
             error1             TYPE c LENGTH 200,
             error2             TYPE c LENGTH 200,
             error3             TYPE c LENGTH 200,
             error4             TYPE c LENGTH 200,
           END OF lty_template.

    DATA et_template      TYPE TABLE OF lty_template.
    DATA et_template_temp TYPE TABLE OF lty_template.
    DATA es_template      TYPE lty_template.

    METHODS Excel_upoad.

ENDCLASS.

CLASS zcl_excel_upload IMPLEMENTATION.
  METHOD excel_upload.
    DATA lv_xstr      TYPE xstring.
// basically it should be called in Behaviour implemetation class where the excel upload is handled
        // LOOP AT entities INTO DATA(ls_entity).
      lv_xstr = ls_entity-file_content.

      TRY.
          DATA(lo_all_sheet) = xco_cp_xlsx=>document->for_file_content( lv_xstr
                )->read_access( )->get_workbook(
                )->worksheet->all.

          DATA(lt_sheets) = lo_all_sheet->get( ).
        CATCH cx_root INTO DATA(lx_root). 
      ENDTRY.
      LOOP AT lt_sheets INTO DATA(ls_sheets) FROM 2.
        DATA(lo_sheet) = xco_cp_xlsx=>document->for_file_content( lv_xstr
      )->read_access( )->get_workbook(
      )->worksheet->for_name( ls_sheets->get_name( ) ).

        DATA(lo_pattern) = xco_cp_xlsx_selection=>pattern_builder->simple_from_to(
             )->from_column( xco_cp_xlsx=>coordinate->for_alphabetic_value( 'A' )
             )->from_row( xco_cp_xlsx=>coordinate->for_numeric_value( 1 )
             )->get_pattern( ).

        lo_sheet->select( lo_pattern
          )->row_stream(
          )->operation->write_to( REF #( lt_template )
          )->set_value_transformation( xco_cp_xlsx_read_access=>value_transformation->string_value
          )->execute( ).
    
      ENDLOOP.
  ENDMETHOD.
ENDCLASS.

