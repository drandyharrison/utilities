/* Based on the blog - How to use SAS to read a range of cells from Excel */
/* https://blogs.sas.com/content/sasdummy/2018/06/21/read-excel-range/ */

/* uses a snapshot of the mpg tracker Excel, to test the functionality */
%LET datadir=/folders/myfolders/sasuser.v94/utilities/;
%LET fname=mpg.xlsx;

/* define footnote for all plots using automatic macro variables */
%LET foot2_arg = color=blue italic;
%LET foot2_val = "Prepared by Andy Harrison [(c) &sysdate9.] at &systime on &sysday";
footnote &foot2_arg &foot2_val;

PROC IMPORT DATAFILE="&datadir.&fname."
	OUT=mpg(RENAME=(Fuel__l_=Fuel VAR5=Price)) DBMS=xlsx REPLACE;
	RANGE="VO62_DVZ$A1:E106";
	INFORMAT Mileage COMMA. Cost NLMNLGBP9.2 Price NLMNLGBP9.3;
	LABEL Fuel="Fuel (l)" Price="Price (£/l)" Cost="Cost (£)";
	FORMAT Cost NLMNLGBP9.1 Price NLMNLGBP9.3 Mileage COMMA9.2;
RUN;

PROC PRINT DATA=mpg LABEL;
RUN;