//Ask user to choose the input and output directories
directory = getDirectory("Choose input directory");
fileList = getFileList(directory);

//Count the maximum number of positions and slices in dataset
run("Bio-Formats Macro Extensions");

newPosition = 0;
nChannel = 0;

//Check all files to ensure identical file structure: even number of series and same number of channels
for (i=0; i<fileList.length; i++) {
	file = directory + fileList[i];
	Ext.setId(file);
	Ext.getSeriesCount(nSeries);
	Ext.getSizeC(sizeChannel)

	//Check to make sure that all files have the same number of channels, otherwise quit with error
	if (sizeChannel != nChannel) {
		if (i < 1) {
			nChannel = sizeChannel;
		}
		//If number of channels differ, give user name of problem file and exit
		else {
			exit("File '" + fileList[i] + "' has a different number of channnels than '" + fileList[i-1] + "'.  Process files individually.");
		}
	}

	//Check to make sure there are at least 2 channels - one autofluor reference channel and one anitbody channel
	if (nChannel < 2){
		exit("File '" + fileList[i] + "' has only one channel.  This macro requires at least an autofluorescence reference channel and an antibody channel.");
	}


	//Check to ensure there is an even number of series in the data set
	if (nSeries%2 != 0){
		exit("File " + fileList[i] + " has an odd number of series.  The file must have an even number of series.");
	}

}

//Ask whether even or odd series are the secondary only controls
seriesSecondary = getNumber("For the first image set, is series 1 or series 2 the secondary only control?", 1); 

//Set other series as the primary series
seriesPrimary = 1;

if (seriesSecondary == 1){
	seriesPrimary = 2;
}


//Ask user which channel is the antibody channel
channelAntibody = getNumber("This dataset contains " + nChannel + " channels.  Which channel is for the antibody staining?", 1);

//Ask user which channel is the antibody channel
channelAutofluor = getNumber("This dataset contains " + nChannel + " channels.  Which channel is for the autofluorescence?", 2);


//Activate batch mode for speed
setBatchMode(true);

//Normalize Autofluor and GFP channel, remove autofluor, save result

//Open all pairs of images
for (i=0; i<fileList.length; i++) {	
	//Set current file name
	file = directory + fileList[i];
	
	//Get number of series for this file
	Ext.setId(file);
	Ext.getSeriesCount(nSeries);
	
	//Open the same position series from each lif file as a hyperstack
	for(a=0; a<nSeries; a+=2) {
		/////////////PROCESS PRIMARY + SECONDARY//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
		//Open the primary antibody series
		run("Bio-Formats Importer", "open=file color_mode=Default view=Hyperstack stack_order=XYCZT series_"+d2s((a+seriesPrimary),0)); 
		
		//Get name of opened stack
		titlePrimary = getTitle();

		//If the series is not an XYZ stack than close image
		getDimensions(width, height, channels, slices, frames);
		if(slices > 1){
			run("Z Project...", "projection=[Max Intensity] all");
			close();	
		}
		
	
		//Split channels and record names of the antibody and autofluor channels
		run("Split Channels");
		channeltitleAntibody = "C" + channelAntibody + "-" + titlePrimary;
		channeltitleAutofluor = "C" + channelAutofluor + "-" + titlePrimary;

		//Measure mean intensity of both channels an calcuate the ratio between the two to get a normalized intensity (normalized to autofluor) 
		selectWindow(channeltitleAntibody);
		getStatistics(dummy, meanAntibody, dummy, dummy, dummy);
		selectWindow(channeltitleAutofluor);
		getStatistics(dummy, meanAutofluor, dummy, dummy, dummy);
		normalizedPrimary = meanAntibody/meanAutofluor;

		//Output results to result table
		setResult("Sample ID", (a/2), titlePrimary);
		setResult("Normalized Primary + Secondary", (a/2), normalizedPrimary);
		
		//Close all images
		close("*");

		/////////////PROCESS SECONDARY ONLY//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
		//Open the primary antibody series
		run("Bio-Formats Importer", "open=file color_mode=Default view=Hyperstack stack_order=XYCZT series_"+d2s((a+seriesSecondary),0)); 
		
		//Get name of opened stack
		titleSecondary = getTitle();

		//If the series is not an XYZ stack than close image
		getDimensions(width, height, channels, slices, frames);
		if(slices > 1){
			run("Z Project...", "projection=[Max Intensity] all");
			close();	
		}
		
	
		//Split channels and record names of the antibody and autofluor channels
		run("Split Channels");
		channeltitleAntibody = "C" + channelAntibody + "-" + titleSecondary;
		channeltitleAutofluor = "C" + channelAutofluor + "-" + titleSecondary;

		//Measure mean intensity of both channels an calcuate the ratio between the two to get a normalized intensity (normalized to autofluor) 
		selectWindow(channeltitleAntibody);
		getStatistics(dummy, meanAntibody, dummy, dummy, dummy);
		selectWindow(channeltitleAutofluor);
		getStatistics(dummy, meanAutofluor, dummy, dummy, dummy);
		normalizedSecondary = meanAntibody/meanAutofluor;

		//Output results to result table
		setResult("Normalized Secondary Only", (a/2), normalizedSecondary);

		//Calulate normalized fold increase of primary+secondary vs. secondary only
		setResult("Ratio Primary+Secondary : Secondary Only", (a/2), normalizedPrimary/normalizedSecondary);

		//Close all images
		close("*");
		
	}
}

setBatchMode(false);

// open contents of Results table in Excel
if (nResults==0) exit("Results table is empty");
pathExcel = directory + "Normalized antibody staining ratio.xls";
saveAs("Results", pathExcel);

exec("open", pathExcel);
exec("cmd", "/c", "start", "excel.exe", pathExcel);
run("Quit"); 

