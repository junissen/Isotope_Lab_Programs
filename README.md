# Isotope_Lab_Programs
Programs for reducing data run on MC-ICP-MS for U/Th dating using SEM and Cups methods

##### Programs include calculation for chemistry and reagent blanks run on SEM methods and standards/age files run on either SEM or Cups methods. 

Programs included are AgeCalculation.py, ChemBlankCalculation.py, ReagenBlankCalculation.py and StandardCalculation.py. Programs can also be accessed as an executable file for either Mac or PC, under the respective folders. To use, download the executable file and to verify proper functioning run *chmod +x "file extension"* in terminal window before running.

Note: A chemblank file must be created using the ChemBlankCalculation program before running AgeCalculation on SEM or Cups. This program uses the output of the chemblank program in its calculuation. Additionally, a spiked standard 234/238 ratio must be calculated using StandardCalculation before running AgeCalculation on Cups. 

Files may be uploaded into the program from anywhere on your computer. Output files will be located in the same folder where your programs are located.

Shown below are the divisions of each file: 

## AgeCalculation.py
---------
Requires sys, Tkinter, tkFileDialog, tkMessageBox, openpyxl, csv, numpy, os, curve_fit and fsolve from scipy.optimize, Figure from matplotlib.figure, FigureCanvasTkAgg from matplotlib.backends.backend_tkagg, islice from itertools, and datetime

* class Application(tk.Frame): Opens Tkinter frame from which the user can choose what age calculation program to run

	* def create_widgets(): choose to change preset values (233 spike conc, sample wt error, spike wt error, 230/232 i and 230/232 i error). If you choose to change them, a secondary window will pop up for you to input values. Once done, click "Submit" and you will be returned to the function method_used(). 
	
	* def method_used(): choose SEM/Cups methods for U run
	
	SEM U run:
	* def upload_sem(): creates manual entry windows for spike used, AS for U run. Prompts whether your AS was the same for your Th run, in case you ran them on different days.
	* def AS_Th_yes(): creates manual entry windows for sample wt, spike wt, sample ID, and row for age spreadsheet, and provides checkbutton option for altering Th file.
	* def AS_Th_no(): prompts what your AS was for your Th run. Creates manual entry windows for sample wt, spike wt, sample ID, and row for age spreadsheet, and provides checkbutton option for altering Th file.
	* def Th_yes_sem(): prompts what row you would like to stop analysis at on Th file. Buttons for uploading Th and Th wash file, and provides checkbutton option of altering U file.
	* def Th_no_sem(): buttons for uploading Th and Th wash file, and provides checkbutton option of altering U file.
	* def U_yes_sem(): prompts what row you would like to stop analysis at on U file. Buttons for uploading U, U wash file, chemblank file, age spreadsheet. Buttons for age calculate and quit.
	* def U_no_sem(): buttons for uploading U, U wash file, chemblank file, age spreadsheet. Buttons for age calculate and quit.
	* def sem_command(): runs age calculate function (Application_sem.age_calculate_sem()) for SEM using uploaded files and input parameters. 
	
	CUPS U run: 
	* def cups_command_U(): choose SEM/Cups methods for Th run
	
		SEM Th run: 
		* def sem_command_Th(): creates manual entry windows for spike used, AS for Th run, sample weight, spike weight, sample ID, and row for age spreadsheet. Provides checkbutton for altering unspiked standard file
		* def unspiked_yes_semcups(): prompts what row you would like to stop analysis at on unspiked file. Buttons for uploading unspiked and unspiked wash standard files. Creates manual entry windows for standard 234/238 ppm value and error value. Provides checkbutton for altering Th file
		* def unspiked_no_semcups(): buttons for uploading unspiked and unspiked wash standard files. Creates manual entry windows for standard 234/238 ppm value and error value. Provides checkbutton for altering Th file
		* def Th_yes_semcups(): prompts what row you would like to stop analysis at on Th file. Buttons for uploading Th and Th wash file, and provides checkbutton option of altering U file. 
		* def Th_no_semcups(): buttons for uploading Th and Th wash file, and provides checkbutton option of altering U file. 
		* def U_yes_semcups(): prompts what row you would like to stop analysis at on U file. Buttons for uploading U, U wash file, chemblank file, age spreadsheets. Prompts whether U wash was run on SEM or Cups method. Buttons for age calculate and quit.
		* def U_no_semcups(): buttons for uploading U, U wash file, chemblank file, age spreadsheets. Prompts whether U wash was run on SEM or Cups method. Buttons for age calculate and quit.
		* def semcups_command(): runs age calculate function (Application_semcups.age_calculation_semcups()) for U on cups and Th on SEM using uploaded files and input parameters.
	
		CUPS Th run:
		* def cups_command_Th(): creates manual entry windows for spike used, sample weight, spike weight, sample ID, and row for age spreadsheet. Provides checkbutton for altering unspiked standard file
		* def unspiked_yes_cups(): prompts what row you would like to stop analysis at on unspiked file. Buttons for uploading unspiked and unspiked wash standard files. Creates manual entry windows for standard 234/238 ppm value and error value. Provides checkbutton for altering Th file
		* def unspiked_no_cups(): buttons for uploading unspiked and unspiked wash standard files. Creates manual entry windows for standard 234/238 ppm value and error value. Provides checkbutton for altering Th file
		* def Th_yes_cups(): prompts what row you would like to stop analysis at on Th file. Buttons for uploading Th and Th wash file, and provides checkbutton option of altering U file. Prompts whether Th wash was run on SEM or Cups method.
		* def Th_no_cups(): buttons for uploading Th and Th wash file, and provides checkbutton option of altering U file. Prompts whether Th wash was run on SEM or Cups method.
		* def U_yes_cups(): prompts what row you would like to stop analysis at on U file. Buttons for uploading U, U wash file, chemblank file, age spreadsheets. Prompts whether U wash was run on SEM or Cups method. Buttons for age calculate and quit.
		* def U_no_cups(): buttons for uploading U, U wash file, chemblank file, age spreadsheets. Prompts whether U wash was run on SEM or Cups method. Buttons for age calculate and quit.
		* def cups_command(): runs age calculate function (Application_cups.age_calculation_cups()) for Cups using uploaded files and input parameters.
	
	WASH OPTIONS FOR CUPS MEASUREMENTS
	* def Uwash_cups(): wash function will be run on Cups method for U
	* def Uwash_sem(): 	wash function will be run on SEM method for U
	* def Thwash_cups(): wash function will be run on Cups method for Th
	* def Thwash_sem(): wash function will be run on SEM method for Th
	
	CHANGING PRESET VALUES
	* def preset_change(): hides main Tkinter window and opens new window. Prompts for changes to 233 spike concentration, sample weight error, spike weight error, 230/232 initial value, and 230/232 initial error value if you would like to change them. To enter values press "Submit" and you will be returned back to the main Tkinter window, with preset values having been changed.
	* def show(): shows the original Tkinter window after changing preset values
	
	PROGRAM QUIT FUNCTION
	* def quit_program(): closes the Tkinter window and ends the Python program 

	UPLOAD FUNCTIONS
	* def file_unspiked_upload_option(): uploads the altered unspiked file. Only .exp files accepted
	* def file_unspiked_upload(): uploads the unspiked file. .exp and .xlsx files accepted
	* def file_unspikedwash_upload(): uploads the unspiked wash file. .exp and .xlsx files accepted
	* def file_Th_upload_option(): uploads the altered Th file. Only .exp files accepted
	* def file_Th_upload(): uploads the Th file. .exp and .xlsx files accepted
	* def file_Thwash_upload(): uploads the Th wash file. .exp and .xlsx files accepted
	* def file_U_upload_option(): uploads the altered U file. Only .exp files accepted
	* def file_U_upload(): uploads the U file. .exp and .xlsx files accepted
	* def file_Uwash_upload(): uploads the U wash file. .exp and .xlsx files accepted
		*all functions above either use given .xlsx file or create temporary .xlsx file for analysis*
	* def file_chemblank_upload(): uploads the chemblank .xlsx file and returns a list of chemblank values
	* def file_upload_export(): uploads the age spreadsheet .xlsx file

AGE CALCULATION FUNCTIONS
	
* class Application_sem(spike, AS, sample wt, spike wt, sample ID, row # on age spreadsheet, Th file, Th wash file, U file, U wash file, chemblank value list, age spreadsheet file, sample weight error, spike weight error, 233 spike conc, 229 spike conc, 230/232i, 230/232i error): class for U and Th on SEM

	* def age_calculate_sem(): calculates and corrects data to determine U/Th age of sample run with U and Th on SEM and exports data into corresponding row in age spreadsheet.
	
* class Application_semcups(spike, AS for Th run, sample wt, spike wt, sample ID, row # on age spreadsheet, spiked standard 234/238, spiked standard 234/238 error, unspiked standard file, unspiked wash standard file, Th file, Th wash file, U file, U wash file, chemblank value list, age spreadsheet file, sample weight error, spike weight error, 233 spike conc, 229 spike conc, 230/232i, 230/232i error, Uwash option (SEM/Cups) ): class for U on cups and Th on SEM 

	* def age_calcualte_semcups(): calculates and corrects data to determine U/Th age of sample run with U on cups and Th on SEM and exports data into corresponding row in age spreadsheet.
	
* class Application_cups(spike, sample wt, spike wt, sample ID, row # on age spreadsheet, spiked standard 234/238, spiked standard 234/238 error, unspiked standard file, unspiked wash standard file, Th file, Th wash file, U file, U wash file, chemblank value list, age spreadsheet file, sample weight error, spike weight error, 233 spike conc, 229 spike conc, 230/232i, 230/232i error, Uwash option (SEM/Cups), Thwash option (SEM/Cups) ): class for U and Th on cups

	* def age_calculate_cups(): calculates and corrects data to determine U/Th age of sample run with U and Th on cups and exports data into corresponding row in age spreadsheet. 
	
### ADDITIONAL FUNCTIONS USED IN AGE CALCULATION
-----
*Accessory functions for cups measurements*
* class unspiked_standard(unspiked standard file, unspiked wash standard file): 

	* def unspiked_calc(): calculates the 237 tail values and errors for the unspiked standard. Returns values as list to be used in tail correction for cups measurements

* class Calculation_forCups(spike, list of unspiked standard tails, list of chemblank values, spike wt, sample wt): corresponding functions of cups measurements

	* def U_calc(U file, U wash file, spiked standard 234/238 value, spiked standard 234/238 value error): completes drift, machine blank, tail, fractionation, spike, chemistry blank, and standard corrections for U ratios on cups measurements. Returns list of corrected values for use in age calculation and Th corrections.
	
	* def Thsem_calc(Th file, Th wash file, list of corrected Th values from SEM filter function): completes machine blank, spike and chemistry blank corrections for imported Th values from SEM function. Returns list of corrected values for use in age calculation. 
	
	* def Thcups_calc(Th file, Th wash file, list of 236/233 corrected ratio and error from U_calc): completes drift, machine blank, tail, fractionation, spike and chemistry blank corrections for Th ratios on cups measurements.

* class isocorrection():  Creates numpy arrays of specified Excel columns and completes element-wise corrections. 
	* def array(filename, column letter): compiles a numpy array from the values of a specified Excel column
	* def drift_correction_offset(source array, ratio array): calculates the offset between the source array value and the calculated value. Compiles numpy array of offset values.  
	* def drift_correction(drift array, source array): corrects the source array for drift, and returns corrected array
	* def drift_correction_alt(drift array, source array, reference array): corrects the source array for drift, and returns corrected array
	* def machine_blank_correction(source array, bottom isotope mean, machine blank mean for bottom isotope, machine blank mean for top isotope): corrects source array for machine blank and returns corrected array
	* def machine_blank_correction_alt(source array, bottom isotope mean, machine blank mean for bottom isotope, machine blank mean for top isotope, ratio array mean): corrects source array for machine blank and returns corrected array
	* def tail_correction(source array, tail 237 top isotope value, tail 237 bottom isotope value, 238/233 machine blank corrected mean, option): corrects source array for 237 tail and returns corrected array. Options are either "norm" (for 234/233, 235/233, and 236/233) or "238/233", as 238/233 tail correction is calculated differently. 
	* def tail_correction_alt(top isotope tail corrected array, bottom isotope tail corrected array): creates new isotope array based off two given tail corrected arrays. Used for calculating 238/235 and 234/238 arrays.
	* def tail_correction_th(source array, 230/229 tail, 232/229 tail, 230/232 tail, 229/232 tail, machine blank ratio, option): corrected the Th source array for 237 tail and returns corrected array. Options are "230/229", "232/229" and "230/232". 
	* def fractionation_correction(source array, ratio array, 236/233 tail corrected mean, top isotope, bottom isotope, spike 236/233 ratio): corrects source array for fractionation and returns corrected array. Only corrects values where the same index value in the ratio array is not NaN. 
----
*Accessory functions for SEM measurements*
* class Ucalculation(spike, AS, U file): 
	
	* def U_normalization_forTh(): filters U SEM measurements and returns list of corrected U ratios for use in Th SEM calculation
	* def U_normalization_forAge(): returns list of corrected U ratios for use in SEM age calculation

* class Thcalculation(spike, AS, Th file, list of corrected U ratios from Ucalculation):
	
	* def Th_normalization_forAge(): filters Th SEM measurements and returns list of corrected Th ratios for use in SEM age calculation
	
* class background_values(U wash file, Th wash file):
	* def U_wash(): calculates and returns list of 233, 234, and 235 wash values in cps
	* def Th_wash(): calculates and returns 230 wash value in cpm

----
* class isofilter(filename, column letter): Calculates unfiltered and filtered mean, standard deviation/error and counts for an Excel column. 

	* def getMean(): calculates the mean of a given column
	* def getStanddev(): calculates the standard deviation of the given column
	* def getCounts(): calculates the total number of cycles in a given column
	* def Filtered_mean(mean, standard deviation, counts, filter number): filters the Excel column based off specific criteria calculated by the mean, standard error, and filter number, and returns the resulting mean. Filter number is 44 for U runs and 28 for Th runs
	* def Filtered_err(mean, standard deviation, counts, filter number): filters the Excel column based off specific criteria calculated by the mean, standard error, and filter number, and returns the resulting 2s error. Filter number is 44 for U runs and 28 for Th runs
	* def Filtered_counts(mean, standard deviation, counts, filter number): filters the Excel column based off specific criteria calculated by the mean, standard error, and filter number, and returns the filtered number of cycles. Filter number is 44 for U runs and 28 for Th runs
    	
* class plot_figure(tk.TK, 234U beam array, 234U index array, 230Th beam array, 230Th index array):

	* def plot_fig(): creates GUI window of 234U and 230Th beam after age calculation to check on beam stability. 
----	
*Accessory functions for changing preset values*

* class Application_preset(tk.Toplevel): creates accessory window if you choose to change preset values

	* def spike_conc_option(): prompts whether you'd like to change the 233 spike concentration. If 233 is changed, 229 later gets calculated dependent on spike used.
	* def spike_yes(): creates manual entry windows for 233 spike concentrations. Prompts whether you'd like to change the sample weight error
	* def spike_no(): prompts whether you'd like to change the sample weight error
	* def samplewt_yes(): creates manual entry window for sample weight error. Prompts whether you'd like to change the spike weight error
	* def samplewt_no(): prompts whether you'd like to change the spike weight error
	* def spikewt_yes(): creates manual entry window for spike weight error. Prompts whether you'd like to change the 230/232 initial value
	* def spikewt_no(): prompts whether you'd like to change the 230/232 initial value
	* def zerotwo_yes(): creates manual entry window for 230/232i and 230/232i error. Provides submit button for finalizing values and returning to main window
	* def zerotwo_no(): provides submit button for finalizing values and returning to main window
	* def click_submit(): edits preset values and returns you to main window
	* def on_closing(): prompts whether you would like to quit program if you "X" out of accessory Tkinter window
----

* def on_closing(): prompts whether you would like to quit program if you "X" out of Tkinter window


### ChemBlankCalculation.py
---------
Requires sys, Tkinter, tkFileDialog, tkMessageBox, numpy, pandas, openpyxl, csv, and os

* class Application(tk.Frame): Opens Tkinter frame from which the user can include machine parameters and upload machine .exp files. 
		
	* def create_widgets(): prompts if you would like to change the spike preset values (233 and 229 concentrations). If yes, a secondary window will open to prompt you.
	* def parameter_input(): creates manual entry windows for blank name, spike info, spike weight, U weight, Th, weight, uptake rate, ionization efficiency, and chemblank export file name. Also provides checkbutton option of altering Th file.
	* def th_yes(): prompts what row you would like to stop analysis at on Th file. Buttons for uploading Th and Th wash file, and provides checkbutton option of altering U file.
	* def th_no(): buttons for uploading Th and Th wash file, and provides checkbutton option of altering U file.
	* def u_yes(): prompts what row you would like to stop analysis at on U file. Buttons for uploading U, U wash file, blank calculate, and quit.
	* def u_no(): buttons for uploading U, U wash file, blank calculate, and quit.
	* def preset_change(): prompts secondary window function to change spike preset values
	* def show(): returns you to original window and continues commands if spike preset values have been changed
	* def quit_program(): closes Tkinter window and ends Python program
	UPLOAD FUNCTIONS
	* def file_upload_th_chemblank(): uploads unaltered Th file and temporarily creates new Th Excel file
	* def file_upload_th_chemblank_option(): uploads altered Th file and temporarily creates new Th Excel file.
	* def file_upload_th_chemblankwash(): uploads Th wash file and temporarily creates new Th wash Excel file
	* def file_upload_u_chemblank(): uploads unaltered U file and temporarily creates new U Excel file
	* def file_upload_u_chemblank_option(): uploads altered U file and temporarily creates new U Excel file
	* def file_upload_u_chemblankwash(): uploads U wash file and temporarily creates new U wash Excel file
	CHEMBLANK RUN
	* def blank_calculate(): Calculates wash and chem blank values for all isotopes. Exports an excel file with isotope data.
------
*Accessory functions for chemblank calculations*

* class chem_blank(): requires filename, columnletter, and isotope analyzed

	* def calc(): calculates the mean, total cycles, and 2s counting statistics error for specified isotope
	
* class plot_figure(tk.TK, 234U beam array, 234U index array, 230Th beam array, 230Th index array):

	* def plot_fig(): creates GUI window of 234U and 230Th beam after age calculation to check on beam stability. 

------
*Accessory functions for changing preset values*

* class Application_preset(tk.Toplevel): creates accessory window if you choose to change preset values

	* def spike_conc_option(): prompts whether you'd like to change the 233 spike concentration. If 233 is changed, 229 later gets calculated dependent on spike used.
	* def spike_yes(): creates manual entry windows for 233 spike concentrations. Provides submit button for finalizing values and returning to main window
	* def spike_no(): provides submit button for finalizing values and returning to main window
	* def click_submit(): edits preset values and returns you to main window
	* def on_closing(): prompts whether you would like to quit program if you "X" out of accessory Tkinter window
----

* def on_closing(): prompts whether you would like to quit program if you "X" out of Tkinter window
	
### ReagentBlankCalculation.py
---------
Requires sys, Tkinter, tkFileDialog, tkMessageBox, numpy, pandas, openpyxl, csv, and os

* class Application(tk.Frame): Opens Tkinter frame from which the user can include machine parameters and upload machine .exp files. 
		
	* def create_widgets():  creates manual entry windows for blank name, solution weight, uptake rate, ionization efficiency, and reagent blank export file name. Also provides checkbutton option of altering Th file.
	* def th_yes(): prompts what row you would like to stop analysis at on Th file. Buttons for uploading Th and Th wash file, and provides checkbutton option of altering U file.
	* def th_no(): buttons for uploading Th and Th wash file, and provides checkbutton option of altering U file.
	* def u_yes(): prompts what row you would like to stop analysis at on U file. Buttons for uploading U, U wash file, blank calculate, and quit.
	* def u_no(): buttons for uploading U, U wash file, blank calculate, and quit.
	* def quit_program(): closes Tkinter window and ends Python program
	UPLOAD FUNCTIONS
	* def file_upload_th_regblank(): uploads unaltered Th file and temporarily creates new Th Excel file
	* def file_upload_th_regblank_option(): uploads altered Th file and temporarily creates new Th Excel file.
	* def file_upload_th_regblankwash(): uploads Th wash file and temporarily creates new Th wash Excel file
	* def file_upload_u_regblank(): uploads unaltered U file and temporarily creates new U Excel file
	* def file_upload_u_regblank_option(): uploads altered U file and temporarily creates new U Excel file
	* def file_upload_u_regblankwash(): uploads U wash file and temporarily creates new U wash Excel file
	REAGENT BLANK RUN
	* def blank_calculate(): Calculates wash and chem blank values for all isotopes. Exports an excel file with isotope data.
REAGENT BLANK ACCESSORY FUNCTION
* class chem_blank(): requires filename, columnletter, and isotope analyzed

	* def calc(): calculates the mean, total cycles, and 2s counting statistics error for specified isotope

* def on_closing(): prompts whether you would like to quit program if you "X" out of Tkinter window


### StandardCalculation.py
--------
Requires sys, Tkinter, tkFileDialog, tkMessageBox, openpyxl, csv, numpy, os, curve_vit from scipy.optimize, Figure from matplotlib.figure, and FigureCanvasTkAgg from matplotlib.backends.backend_tkagg.

* class Application(tk.Frame): Opens master Tkinter frame from which the users are directed to either a SEM or Cups standard run.
	* def create_widgets(): creates checkbutton option for either SEM or Cups standard run
	* def sem_command(): prompts the class Application_sem() to run if the SEM checkbutton is marked
	* def cups_command(): prompts the class Application_cups() to run if the Cups checkbutton is marked
	
SEM STANDARD
* class Application_sem(tk.Frame): Adds SEM file uploads and machine parameters to master Tkinter frame, and prompts SEM standard calculations. 
	* def create_widgets_sem(): creates manual entry windows for AS, 234U wash, and spike information. Provides prompt for altering standard file.
	* def option_no():  provides upload button for unaltered 112A standard file, standard calculation, and quit
	* def option_yes(): prompts what row you would like to stop analysis at on 112A file. Provides upload button for altered 112A standard file, standard calculation, and quit
	* def quit_program(): exits Tkinter window and ends Python program
	UPLOAD FUNCTIONS
	* def file_usem_upload(): uploads unaltered 112A standard file and temporarily creates new 112A standard Excel file
	* def file_usem_upload_option(): uploads altered 112A standard file and temporarily creates new 112A standard Excel file
	SEM STANDARD RUN
	* def standard(): completes standard calculations for SEM, and results in a message box including the 236/233, 235/233 and d234 values for your standard run. Also results in display of 234U beam intensity. 

CUPS STANDARD   
* class Application_cups(tk.Frame): Adds Cups file uploads and machine parameters to master Tkinter frame, and prompts Cups standard calculations.
	* def create_widgets_cups(): creates manual entry window for spike information. Provides prompt for altering unspiked file.
	* def unspiked_yes(): prompts what row you would like to stop analysis at on unspiked file. Provides upload buttons for altered unspiked file and wash. Provides prompt for altering spiked file
	* def unspiked_no(): Provides upload buttons for unaltered unspiked file and wash. Provides prompt for altering spiked file
	* def spiked_yes(): prompts what row you would like to stop analysis at on spiked file. Provides upload buttons for altered spiked file and wash, standard calculation, and quit. Prompts whether standard wash run on SEM or Cups method. 
	* def spiked_no(): Provides upload buttons for unaltered spiked file and wash, standard calculation, and quit. Prompts whether standard wash run on SEM or Cups method. 
	* def quit_program(): exits Tkinter window and ends Python program
	WASH OPTIONS
	* def Uwash_cups(): sets wash to Cups method for spiked standard wash run
	* def Uwash_sem(): sets wash to SEM method for spiked standard wash run
	UPLOAD FUNCTIONS
	* def file_unspiked_upload(): uploads the unaltered unspiked file and temporarily creates new unspiked Excel file
	* def file_unspiked_upload_option(): uploads the altered unspiked file and temporarily creates new unspiked Excel file
	* def file_unspiked_wash_upload(): uploads the unspiked wash file and temporarily creates new unspiked wash Excel file
	* def file_spiked_upload(): uploads the spiked file and temporarily creates new spiked Excel file
	* def file_spiked_wash_upload(): uploads the spiked wash file and temporarily creates new spiked wash Excel file
	CUPS STANDARD RUN
	* def standard(): runs the unspiked standard function, followed by the spiked standard function
	* def unspiked_standard(): calculates the 237 tail values and errors for the unspiked standard. These are then use in tail correction for the spiked standard
	* def spiked_standard(): completes the standard calculations for Cups, and results in a message box including the 234/238, 237/238, 236/233, 238/235 and d234 values for your standard run. Also results in display of 234U beam intensity. 
   
ADDITIONAL FUNCTIONS FOR STANDARD CALCULATION
* class plot_figure(tk.Tk): Creates new Toplevel widget for displaying plot

	* def plot_234(filename, column1, column2): creates a plot of cycle number versus beam intensity. Column 1 denotes x data, column 2 denotes y data.
    
* class isofilter(filename, column letter, filter number): Calculates unfiltered and filtered mean, standard deviation/error and counts for an Excel column. Filter number is 44 for U runs. 
	
	* def getMean(): calculates the mean of the given column
	* def getStanddev(): calculates the standard deviation of the given column
	* def getCounts(): calculates the total number of cycles in a given column
	* def Filtered_mean(mean, standard deviation, counts): filters the Excel column based off specific criteria calculated by the mean, standard error, and filter number, and returns the resulting mean.
	* def Filtered_err(mean, standard deviation, counts): filters the Excel column based off specific criteria calculated by the mean, standard error, and filter number, and returns the resulting 2s error. 
	* def Filtered_counts(mean, standard deviation, counts): filters the Excel column based off specific criteria calculated by the mean, standard error, and filter number, and returns the filtered number of cycles. 
    
* class isocorrection():  Creates numpy arrays of specified Excel columns and completes element-wise corrections.

	* def array(filename, column letter): compiles a numpy array from the values of a specified Excel column
	* def drift_correction_offset(source array, ratio array): calculates the offset between the source array value and the calculated value. Compiles numpy array of offset values. 
	* def drift_correction(drift array, source array): corrects the source array for drift, and returns corrected array
	* def machine_blank_correction(source array, bottom isotope mean, machine blank mean for bottom isotope, machine blank mean for top isotope): corrects source array for machine blank and returns corrected array
	* def tail_correction(source array, tail 237 top isotope value, tail 237 bottom isotope value, 238/233 machine blank corrected mean, option): corrects source array for 237 tail and returns corrected array. Options are either "norm" (for 234/233, 235/233, and 236/233) or "238/233", as 238/233 tail correction is calculated differently. 
	* def tail_correction_alt(top isotope tail corrected array, bottom isotope tail corrected array): creates new isotope array based off two given tail corrected arrays. Used for calculating 238/235 and 234/238 arrays.
	* def fractionation_correction(source array, ratio array, 236/233 tail corrected mean, top isotope, bottom isotope, spike 236/233 ratio): corrects source array for fractionation and returns corrected array. Only corrects values where the same index value in the ratio array is not NaN. 
    
* def on_closing(): prompts whether you would like to quit program if you "X" out of Tkinter window
    
    




 

 










