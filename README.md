# utl-using-sql-instead-the-excel-formula-language-for-solving-excel-problems-pyodbc
Creating skilled excel programmers out out of skilled ansi sql programmers
    %let pgm=utl-using-sql-instead-the-excel-formula-language-for-solving-excel-problems-pyodbc;

    Creating skilled excel programmers out out of skilled ansi sql programmers

    All the processing is done inside excel no R or Python dataframe is created,
    This code should run everywhere ansi sql is available?

    ONLY PYODBC PROPERLY SUPPORTS SQL 'CREATE TABLE' INSIDE EXCEL

     I Python: XLWINGS, OPENXLSX, PYXLL, XLSWRITE, XLRD, XLWT
     2 R     : RODBC, ODBS, OPENXLS, XLCONNECT, XLSX, RODBCDBI

     FOUR SOLUTIONS  (sheet names and named ranges can be used interchangeably)

       1. join sheets:   create new sheet(join_sheets) by concatenating and summarizing sheets
       2. join ranges    create new sheet(join_ranges) and range(join_ranges) by concatenating and summarizing sheets

       3. concat sheets: create sheet(concat_sheets)  by concating sheet1 and sheet2 a 
                         This works with sheet name or named ranges

       4. inline data:   creating and populating a sheet

     Python package pyodbc seems more powerfull than r packages ODBC, RODBC or RODBCDBI.
     In particulat I wanted the sql 'CREATE TABLE' statement to maintain pure sql.

    All code must be left justified before submission.

    github
    https://tinyurl.com/5x7nfufs
    https://github.com/rogerjdeangelis/utl-using-sql-instead-the-excel-formula-language-for-solving-excel-problems-pyodbc


    /*****************************************************************************************************************************************/
    /*                                                  |                                                                                    */
    /*                 INPUT                            |                           PROCESS                                                  */
    /*                                                  |                                                                                    */
    /*     INPUT TWO SHEETS WITH NAMED RANGES           |                                                                                    */
    /*                                                  |  1. JOIN SHEETS                                                                    */
    /* d:/xls/theesheets.xlsx named range=sheet1        |                                                                                    */
    /*  named range males = sheet2!$A$1:$C#4            |  %inc "c:/utl/utl_mkeodbc.sas" / nosource;                                         */
    /*                                                  |                                                                                    */
    /*      +---------------+-----+                     |  %utl_pybegin;                                                                     */
    /*      |  A    |  B    |  C  |  MALE STUDENTS      |  parmcards4;                                                                       */
    /*      +---------------+-----+                     |  import pyodbc                                                                     */
    /*  1   |NAME   |SEX    |AGE  |                     |  import pandas as pd                                                               */
    /*      |-------+-------|-----|                     |  conn_str = (                                                                      */
    /*  2   |Alfred |M      |13   |                     |      r'Driver={Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb)};'           */
    /*      |-------+-------+-----+                     |      r'DBQ=d:/xls/have.xlsx;'                                                      */
    /*  3   |Alex   |M      |14   |                     |      r'ReadOnly=0;'                                                                */
    /*      |-------+-------+-----+                     |  );                                                                                */
    /*  4   |JAMES  |M      |15   |                     |  conn = pyodbc.connect(conn_str, autocommit=True)                                  */
    /*      -----------------------                     |  cursor = conn.cursor()                                                            */
    /*      [SHEET]                                     |  cursor.execute("create table join_sheets (sex varchar(255), mean_age numeric)")   */
    /*                                                  |  query = """                                                                       */
    /*   named range females = sheet2!$A$1:$C#4         |    insert into join_sheets                                                         */
    /*                                                  |    select                            OUTPUT (NEW THIRD SHEET MEAN AGE BY SEX)      */
    /*      +---------------------+                     |       sex                                                                          */
    /*      |  A  |  B    |  C    |  FEMALE STUDENTS    |     , avg(age) as mean_age           d:/xls/have.xlsx                              */
    /*      +---------------------+                     |    from                                                                            */
    /*  1   |NAME   |SEX  |AGE    |                     |       (select                        join_sheets=join_sheets!$A$1:$C#3             */
    /*      |-------+-----|-------|                     |           *                                                                        */
    /*  2   |Alice  |F    |12     |                     |        from                          +-------------+-----+                         */
    /*      |-------+-----+-------+                     |           [sheet1$]                  |  A    |  B  |MEAN |                         */
    /*  3   |Barbara|F    |13     |                     |        union                         +-------------+-----+                         */
    /*      |-------+-----+-------+                     |        select                      1 |ROWS   |SEX  |AGE  |                         */
    /*  4   |Carol  |F    |14     |                     |           *                          |-------+-----|-----|                         */
    /*      -----------------------                     |        from                        2 | 1     |M    |14   |                         */
    /*      [SHEET2]                                    |           [sheet2$])                 |-------+-----+-----+                         */
    /*                                                  |        group                       3 | 2     |F    |13   |                         */
    /*   CREATE INPUT WORKBOOK                          |           by sex                     ---------------------                         */
    /*                                                  |  """                                 [JOIN_SHEETS]                                 */
    /*   %utlfkil(d:/xls/have.xlsx);                    |  cursor.execute(query)                                                             */
    /*                                                  |  conn.commit()                                                                     */
    /*   %utl_rbegin;                                   |  conn.close()                                                                      */
    /*   parmcards4;                                    |  ;;;;                                                                              */
    /*   library(openxlsx);                             |  %utl_pyend;                                                                       */
    /*   males <- read.table(header = TRUE, text = "    |                                                                                    */
    /*   NAME   SEX AGE                                 |------------------------------------------------------------------------------------*/
    /*   Alfred Male   13                               |                                                                                    */
    /*   Alex   Male   14                               |  2. CONCATENATE SHEETS                                                             */
    /*   James  Male   15                               |                                                                                    */
    /*   ");                                            |  %inc "c:/utl/utl_mkeodbc.sas" / nosource; creates input                           */
    /*   males;                                         |                                                                                    */
    /*   females <- read.table(header = TRUE, text = "  |  %utl_pybegin;                                                                     */
    /*   NAME   SEX AGE                                 |  parmcards4;                                                                       */
    /*   Alice   Female 12                              |  import pyodbc                                                                     */
    /*   Barbara Female 13                              |  import pandas as pd                                                               */
    /*   Carol   Female 14                              |  conn_str = (                                                                      */
    /*   ");                                            |   r'Driver={Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb)};'              */
    /*   females;                                       |   r'DBQ=d:/xls/have.xlsx;'                                                         */
    /*   library(openxlsx);                             |   r'ReadOnly=0;'                                                                   */
    /*   wb <- createWorkbook();                        |  );                                                                                */
    /*   addWorksheet(wb, "sheet1")                     |  conn = pyodbc.connect(conn_str, autocommit=True)                                  */
    /*   addWorksheet(wb, "sheet2")                     |  cursor = conn.cursor()                                                            */
    /*   writeDataTable(wb,"sheet1",x=males);           |  cursor.execute("CREATE TABLE concat_ranges \                                      */
    /*   writeDataTable(wb,"sheet2",x=females);         |     (name varchar(255), sex varchar(255), age numeric)")                           */
    /*   createNamedRegion(                             |  query = """                                                                       */
    /*     wb = wb,                                     |    insert into concat_ranges                                                       */
    /*     sheet = 1,                                   |    select                           OUTPUT (COMBINED SHEETS)                       */
    /*     name = "males",                              |       name                                                                         */
    /*     rows = 1:(nrow(males) + 1),                  |      ,sex                          +-------------------+                           */
    /*     cols = 1:ncol(males)                         |      ,age                          |  A    |  B  |  C  |                           */
    /*   );                                             |    from                            +-------------+-----+                           */
    /*   createNamedRegion(                             |       ( select                   1 |NAME   |SEX  |AGE  |                           */
    /*     wb = wb,                                     |           name                     |-------+-----|-----|                           */
    /*     sheet = 2,                                   |          ,sex                    2 |Alfred |M    |13   |                           */
    /*     name = "females",                            |          ,age                      |-------+-----+-----+                           */
    /*     rows = 1:(nrow(females) + 1),                |        from                      3 |Alex   |M    |14   |                           */
    /*     cols = 1:ncol(females)                       |           males                    |-------+-----+-----+                           */
    /*   );                                             |        union                     4 |James  |M    |15   |                           */
    /*   saveWorkbook(wb, file = "d:/xls/have.xlsx"     |        select                      |-------+-----|-----|                           */
    /*    , overwrite = TRUE);                          |            name                  5 |Alice  |F    |12   |                           */
    /*   ;;;;                                           |           ,sex                     |-------+-----+-----+                           */
    /*   %utl_rend;                                     |           ,age                   6 |Barbara|F    |13   |                           */
    /*                                                  |        from                        |-------+-----+-----+                           */
    /*                                                  |           females )              7 |Carol  |F    |14   |                           */
    /*                                                  |  """                               ---------------------                           */
    /*                                                  |  cursor.execute(query)             [CONCAT_SHEETS]                                 */
    /*                                                  |  conn.commit()                                                                     */
    /*                                                  |  conn.close()                                                                      */
    /*                                                  |  ;;;;                                                                              */
    /*                                                  |  %utl_pyend;                                                                       */
    /*                                                  |                                                                                    */
    /*                                                  |------------------------------------------------------------------------------------*/
    /*                                                  |                                                                                    */
    /*                                                  |  4. CREATE SHEET USING INLINE DATA                                                 */
    /*                                                  |                                                                                    */
    /*                                                  |  %inc "c:/utl/utl_mkeodbc.sas" / nosource;                                         */
    /*                                                  |                                                                                    */
    /*                                                  |  %utl_pybegin;                                                                     */
    /*                                                  |  parmcards4;                                                                       */
    /*                                                  |  import pyodbc                                                                     */
    /*                                                  |  conn_str = (                                                                      */
    /*                                                  |   r'Driver={Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb)};'              */
    /*                                                  |   r'DBQ=d:/xls/have.xlsx;'                                                         */
    /*                                                  |   r'ReadOnly=0;'                                                                   */
    /*                                                  |  );                                                                                */
    /*                                                  |  conn = pyodbc.connect(conn_str, autocommit=True)            OUTPUT (CREATE SHEET) */
    /*                                                  |  cursor = conn.cursor()                                                            */
    /*                                                  |  cursor.execute("""CREATE TABLE pop_sheet                    +-------------------+ */
    /*                                                  |    (name varchar(255), sex varchar(255), age numeric)""")    |  A    |  B  |  C  | */
    /*                                                  |  cursor.execute("""INSERT INTO pop_sheet                     +-------------+-----+ */
    /*                                                  |    (name, sex, age) values (?, ?, ?)""", 'Roger', 'M', 14) 1 |NAME   |SEX  |AGE  | */
    /*                                                  |  cursor.execute("""INSERT INTO pop_sheet                     |-------+-----|-----| */
    /*                                                  |    (name, sex, age) values (?, ?, ?)""", 'Alexi', 'F', 14) 2 |Rogerd |M    |14   | */
    /*                                                  |  conn.commit()                                               |-------+-----+-----+ */
    /*                                                  |  conn.close()                                              3 |Alexi  |F    |14   | */
    /*                                                  |  ;;;;                                                        |-------+-----+-----+ */
    /*                                                  |  %utl_pyend;                                                 [pop_sheet]           */
    /*                                                  |                                                                                    */
    /*                                                  |                                                                                    */
    /*****************************************************************************************************************************************/

    /*                   _
    (_)_ __  _ __  _   _| |_
    | | `_ \| `_ \| | | | __|
    | | | | | |_) | |_| | |_
    |_|_| |_| .__/ \__,_|\__|
            |_|
    */

    %utlfkil(d:/xls/have.xlsx);

    %utl_rbegin;
    parmcards4;
    library(openxlsx);
    males <- read.table(header = TRUE, text = "
    NAME   SEX AGE
    Alfred Male   13
    Alex   Male   14
    James  Male   15
    ");
    males;
    females <- read.table(header = TRUE, text = "
    NAME   SEX AGE
    Alice   Female 12
    Barbara Female 13
    Carol   Female 14
    ");
    females;
    library(openxlsx);
    wb <- createWorkbook();
    addWorksheet(wb, "sheet1")
    addWorksheet(wb, "sheet2")
    writeDataTable(wb,"sheet1",x=males);
    writeDataTable(wb,"sheet2",x=females);
    createNamedRegion(
      wb = wb,
      sheet = 1,
      name = "males",
      rows = 1:(nrow(males) + 1),
      cols = 1:ncol(males)
    );
    createNamedRegion(
      wb = wb,
      sheet = 2,
      name = "females",
      rows = 1:(nrow(females) + 1),
      cols = 1:ncol(females)
    );
    saveWorkbook(wb, file = "d:/xls/have.xlsx"
     , overwrite = TRUE);
    ;;;;
    %utl_rend;

    /**************************************************************************************************************************/
    /*                                                                                                                        */
    /*      INPUT TWO SHEETS                                                                                                  */
    /*                                                                                                                        */
    /*  d:/xls/theesheets.xlsx named range=sheet1                                                                             */
    /*   named range males = sheet2!$A$1:$C#4                                                                                 */
    /*                                                                                                                        */
    /*       +---------------+-----+                                                                                          */
    /*       |  A    |  B    |  C  |                                                                                          */
    /*       +---------------+-----+                                                                                          */
    /*   1   |NAME   |SEX    |AGE  |                                                                                          */
    /*       |-------+-------|-----|                                                                                          */
    /*   2   |Alfred |M      |13   |                                                                                          */
    /*       |-------+-------+-----+                                                                                          */
    /*   3   |Alex   |M      |14   |                                                                                          */
    /*       |-------+-------+-----+                                                                                          */
    /*   4   |JAMES  |M      |15   |                                                                                          */
    /*       -----------------------                                                                                          */
    /*       [SHEET]                                                                                                          */
    /*                                                                                                                        */
    /*    named range females = sheet2!$A$1:$C#4                                                                              */
    /*                                                                                                                        */
    /*       +---------------------+                                                                                          */
    /*       |  A  |  B    |  C    |                                                                                          */
    /*       +---------------------+                                                                                          */
    /*   1   |NAME   |SEX  |AGE    |                                                                                          */
    /*       |-------+-----|-------|                                                                                          */
    /*   2   |Alice  |F    |12     |                                                                                          */
    /*       |-------+-----+-------+                                                                                          */
    /*   3   |Barbara|F    |13     |                                                                                          */
    /*       |-------+-----+-------+                                                                                          */
    /*   4   |Carol  |F    |14     |                                                                                          */
    /*       -----------------------                                                                                          */
    /*       [SHEET2]                                                                                                         */
    /*                                                                                                                        */
    /**************************************************************************************************************************/

    /*     _       _             _               _
    / |   (_) ___ (_)_ __    ___| |__   ___  ___| |_ ___
    | |   | |/ _ \| | `_ \  / __| `_ \ / _ \/ _ \ __/ __|
    | |   | | (_) | | | | | \__ \ | | |  __/  __/ |_\__ \
    |_|  _/ |\___/|_|_| |_| |___/_| |_|\___|\___|\__|___/
        |__/
    */

    /*----                                                                   ----*/
    /*---- this create the input workbook                                    ----*/
    /*----                                                                   ----*/

    %inc "c:/utl/utl_mkeodbc.sas" / nosource; /*---- sources input program   ----*/

    %utl_pybegin;
    parmcards4;
    import pyodbc
    import pandas as pd
    conn_str = (
        r'Driver={Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb)};'
        r'DBQ=d:/xls/have.xlsx;'
        r'ReadOnly=0;'
    );
    conn = pyodbc.connect(conn_str, autocommit=True)
    cursor = conn.cursor()
    cursor.execute("create table join_sheets (sex varchar(255), mean_age numeric)")
    query = """
      insert into join_sheets
      select
         sex
       , avg(age) as mean_age
      from
         (select
             *
          from
             [sheet1$]
          union
          select
             *
          from
             [sheet2$])
          group
             by sex
    """
    cursor.execute(query)
    conn.commit()
    conn.close()
    ;;;;
    %utl_pyend;

    /**************************************************************************************************************************/
    /*                                                                                                                        */
    /*      OUTPUT (NEW THIRD SHEET)                                                                                          */
    /*                                                                                                                        */
    /*      d:/xls/have.xlsx                                                                                                  */
    /*                                                                                                                        */
    /*      join_sheets=join_sheets!$A$1:$C#3                                                                                */
    /*                                                                                                                        */
    /*      +---------------+-----+                                                                                           */
    /*      |  A    |  B    |MEAN |                                                                                           */
    /*      +---------------+-----+                                                                                           */
    /*    1 |ROWS   |SEX    |AGE  |                                                                                           */
    /*      |-------+-------|-----|                                                                                           */
    /*    2 | 1     |M      |14   |                                                                                           */
    /*      |-------+-------+-----+                                                                                           */
    /*    3 | 2     |F      |13   |                                                                                           */
    /*      -----------------------                                                                                           */
    /*      [JOIN_SHEETS]                                                                                                     */
    /*                                                                                                                        */
    /**************************************************************************************************************************/

    /*___      _       _
    |___ \    (_) ___ (_)_ __    _ __ __ _ _ __   __ _  ___  ___
      __) |   | |/ _ \| | `_ \  | `__/ _` | `_ \ / _` |/ _ \/ __|
     / __/    | | (_) | | | | | | | | (_| | | | | (_| |  __/\__ \
    |_____|  _/ |\___/|_|_| |_| |_|  \__,_|_| |_|\__, |\___||___/
            |__/                                 |___/
    */

    %inc "c:/utl/utl_mkeodbc.sas" / nosource;

    %utl_pybegin;
    parmcards4;
    import pyodbc
    import pandas as pd
    conn_str = (
     r'Driver={Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb)};'
     r'DBQ=d:/xls/have.xlsx;'
     r'ReadOnly=0;'
    );
    conn = pyodbc.connect(conn_str, autocommit=True)
    cursor = conn.cursor()
    cursor.execute("CREATE TABLE join_ranges (sex varchar(255), mean_age numeric)")
    query = """
      insert into join_ranges
      select
         sex
       , avg(age) as mean_age
      from
         (select
             *
          from
             males
          union
          select
             *
          from
             females)
          group
             by sex
    """
    cursor.execute(query)
    conn.commit()
    conn.close()
    ;;;;
    %utl_pyend;


    /**************************************************************************************************************************/
    /*                                                                                                                        */
    /*      OUTPUT (NEW THIRD SHEET)                                                                                          */
    /*                                                                                                                        */
    /*      d:/xls/have.xlsx                                                                                                  */
    /*                                                                                                                        */
    /*      join_sheets=join_ranges!$A$1:$C#3                                                                                 */
    /*                                                                                                                        */
    /*      +---------------+-----+                                                                                           */
    /*      |  A    |  B    |MEAN |                                                                                           */
    /*      +---------------+-----+                                                                                           */
    /*    1 |ROWS   |SEX    |AGE  |                                                                                           */
    /*      |-------+-------|-----|                                                                                           */
    /*    2 | 1     |M      |14   |                                                                                           */
    /*      |-------+-------+-----+                                                                                           */
    /*    3 | 2     |F      |13   |                                                                                           */
    /*      -----------------------                                                                                           */
    /*      [JOIN_RANGES]                                                                                                     */
    /*                                                                                                                        */
    /**************************************************************************************************************************/
    /*____                             _         _               _
    |___ /    ___ ___  _ __   ___ __ _| |_   ___| |__   ___  ___| |_ ___
      |_ \   / __/ _ \| `_ \ / __/ _` | __| / __| `_ \ / _ \/ _ \ __/ __|
     ___) | | (_| (_) | | | | (_| (_| | |_  \__ \ | | |  __/  __/ |_\__ \
    |____/   \___\___/|_| |_|\___\__,_|\__| |___/_| |_|\___|\___|\__|___/

    */

    %inc "c:/utl/utl_mkeodbc.sas" / nosource; creates input

    %utl_pybegin;
    parmcards4;
    import pyodbc
    import pandas as pd
    conn_str = (
     r'Driver={Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb)};'
     r'DBQ=d:/xls/have.xlsx;'
     r'ReadOnly=0;'
    );
    conn = pyodbc.connect(conn_str, autocommit=True)
    cursor = conn.cursor()
    cursor.execute("CREATE TABLE concat_ranges \
       (name varchar(255), sex varchar(255), age numeric)")
    query = """
      insert into concat_ranges
      select
         name
        ,sex
        ,age
      from
         ( select
             name
            ,sex
            ,age
          from
             males
          union
          select
              name
             ,sex
             ,age
          from
             females )
    """
    cursor.execute(query)
    conn.commit()
    conn.close()
    ;;;;
    %utl_pyend;

    /**************************************************************************************************************************/
    /*                                                                                                                        */
    /*                                                                                                                        */
    /*     OUTPUT (COMBINED SHEETS)                                                                                           */
    /*                                                                                                                        */
    /*    +-------------------+                                                                                               */
    /*    |  A    |  B  |  C  |                                                                                               */
    /*    +-------------+-----+                                                                                               */
    /*  1 |NAME   |SEX  |AGE  |                                                                                               */
    /*    |-------+-----|-----|                                                                                               */
    /*  2 |Alfred |M    |13   |                                                                                               */
    /*    |-------+-----+-----+                                                                                               */
    /*  3 |Alex   |M    |14   |                                                                                               */
    /*    |-------+-----+-----+                                                                                               */
    /*  4 |James  |M    |15   |                                                                                               */
    /*    |-------+-----|-----|                                                                                               */
    /*  5 |Alice  |F    |12   |                                                                                               */
    /*    |-------+-----+-----+                                                                                               */
    /*  6 |Barbara|F    |13   |                                                                                               */
    /*    |-------+-----+-----+                                                                                               */
    /*  7 |Carol  |F    |14   |                                                                                               */
    /*    ---------------------                                                                                               */
    /*    [CONCAT_SHEETS]                                                                                                     */
    /*                                                                                                                        */
    /**************************************************************************************************************************/

    /*  _     _       _ _                  _       _
    | || |   (_)_ __ | (_)_ __   ___    __| | __ _| |_ __ _
    | || |_  | | `_ \| | | `_ \ / _ \  / _` |/ _` | __/ _` |
    |__   _| | | | | | | | | | |  __/ | (_| | (_| | || (_| |
       |_|   |_|_| |_|_|_|_| |_|\___|  \__,_|\__,_|\__\__,_|

    */

    %inc "c:/utl/utl_mkeodbc.sas" / nosource;

    %utl_pybegin;
    parmcards4;
    import pyodbc
    conn_str = (
     r'Driver={Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb)};'
     r'DBQ=d:/xls/have.xlsx;'
     r'ReadOnly=0;'
    );
    conn = pyodbc.connect(conn_str, autocommit=True)
    cursor = conn.cursor()
    cursor.execute("""CREATE TABLE pop_sheet
      (name varchar(255), sex varchar(255), age numeric)""")
    cursor.execute("""INSERT INTO pop_sheet
      (name, sex, age) values (?, ?, ?)""", 'Roger', 'M', 14)
    cursor.execute("""INSERT INTO pop_sheet
      (name, sex, age) values (?, ?, ?)""", 'Alexi', 'F', 14)
    conn.commit()
    conn.close()
    ;;;;
    %utl_pyend;

    /**************************************************************************************************************************/
    /*                                                                                                                        */
    /*     OUTPUT (CREATE AND POPULATE SHEET)                                                                                 */
    /*                                                                                                                        */
    /*    +-------------------+                                                                                               */
    /*    |  A    |  B  |  C  |                                                                                               */
    /*    +-------------+-----+                                                                                               */
    /*  1 |NAME   |SEX  |AGE  |                                                                                               */
    /*    |-------+-----|-----|                                                                                               */
    /*  2 |Rogerd |M    |14   |                                                                                               */
    /*    |-------+-----+-----+                                                                                               */
    /*  3 |Alexi  |F    |14   |                                                                                               */
    /*    |-------+-----+-----+                                                                                               */
    /*    [pop_sheet]                                                                                                         */
    /*                                                                                                                        */
    /**************************************************************************************************************************/



    /*              _
      ___ _ __   __| |
     / _ \ `_ \ / _` |
    |  __/ | | | (_| |
     \___|_| |_|\__,_|

    */
