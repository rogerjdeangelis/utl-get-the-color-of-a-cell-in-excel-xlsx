# utl-get-the-color-of-a-cell-in-excel-xlsx
Get the color of a cell in excel xlsx
    Get the color of a cell in excel xlsx

     Process
       a. Use proc report and ods excel to create an excel cell with
          background rgb color "CXEE0044", torch red
           call define (_col_, "STYLE", "style=[backgroundcolor=CXEE0044");
       b. Use python openpyxl package to read the color


    Githib
    https://tinyurl.com/4258yw8k
    https://github.com/rogerjdeangelis/utl-get-the-color-of-a-cell-in-excel-xlsx


    StackOverflow
    https://tinyurl.com/4eudj2ze
    https://stackoverflow.com/questions/55122922/get-the-color-of-a-cell-from-xlsx-with-python

    Sumit Pokhrel
    https://stackoverflow.com/users/2690723/sumit-pokhrel

    *_                   _
    (_)_ __  _ __  _   _| |_
    | | '_ \| '_ \| | | | __|
    | | | | | |_) | |_| | |_
    |_|_| |_| .__/ \__,_|\__|
            |_|
    ;

    * CREATE SHEET=CLASS WITH RGB COLOR 'TORCH RED' "EE0044" in CELL A2;

    %utlfkil(d:/xls/have.xlsx);

    ods excel file="d:/xls/have.xlsx" options(sheet_name="class");;
    proc report data=sashelp.class(obs=1);
    cols name age sex;
    define name / dislay;
    compute name;
      call define (_col_, "STYLE", "style=[backgroundcolor=CXEE0044");
    endcomp;
    run;quit;
    ods excel close;


    * JOYCE HAS BACKGROUND COLLOR OF TORCH RED #EE0044;

    d:/xls/have.xlsx

    NAME      AGE      SEX
    Joyce      11      F

    *
     _ __  _ __ ___   ___ ___  ___ ___
    | '_ \| '__/ _ \ / __/ _ \/ __/ __|
    | |_) | | | (_) | (_|  __/\__ \__ \
    | .__/|_|  \___/ \___\___||___/___/
    |_|
    ;

    * BEST;

    * incase you rerun;
    %utlfkil(d:/xpt/want.xpt);

    proc datasets lib=work nolist;
     delete want;
    run;quit;

    %utl_submit_py64_38("
    import numpy as np;
    import pandas as pd;
    import pyreadstat;
    import openpyxl;
    from openpyxl import load_workbook;
    excel_file = 'd:/xls/have.xlsx';
    wb = load_workbook(excel_file, data_only = True);
    sh = wb['class'];
    color_in_hex = sh['A2'].fill.start_color.index;
    rgb=pd.DataFrame(list(color_in_hex[2:]), columns=['rgb']).T;
    adx=rgb.iloc[0,0] + rgb.iloc[0,1] + rgb.iloc[0,2] + rgb.iloc[0,3] + rgb.iloc[0,4] + rgb.iloc[0,5];
    rgb['col']=adx;
    want=pd.DataFrame(rgb['col']);
    pyreadstat.write_xport(want,'d:/xpt/want.xpt',table_name='want');
    ");

    *            _               _
      ___  _   _| |_ _ __  _   _| |_
     / _ \| | | | __| '_ \| | | | __|
    | (_) | |_| | |_| |_) | |_| | |_
     \___/ \__,_|\__| .__/ \__,_|\__|
                    |_|
    ;

    libname xpt xport "d:/xpt/want.xpt";

    proc print data=xpt.want;
    run;quit;

    libname xpt clear;

    Obs     COL

      1    EE0044



