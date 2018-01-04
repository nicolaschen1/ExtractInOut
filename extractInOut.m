%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
% extractInOut.m
% 
% Description: Extract the list of the inports and outports 
% for a subsystem in an excel file
%
% Author: Nicolas Chen
%
% Prerequisities: open the Simulink model, select the subsystem then run this script.
%
% Inputs: Simulink model (*.mdl or *.slx), name of the excel file
%
% Output: Excel file filled with inports and outports.
%
% VERSION: 1.0
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

pwd;
pathRep = pwd;

try
    line = find_system(gcs, 'SearchDepth', 1, 'FindAll', 'on','Selected', 'on');
    if (size(line,1) == 1)       
        nameSubsystem = getfullname(line);
        errorSubsystem = 'FALSE';        
    else
        errorSubsystem = 'TRUE';
        errordlg('Subsystem Name is not at the root','Subsystem Error');
    end
catch
    errordlg('Subsystem Name not found','Subsystem Error');
    errorSubsystem = 'TRUE';
end

if (strcmp(errorSubsystem,'TRUE'))
    errorNameExcelFile = 'TRUE';
else

%dialogbox
prompt = {'ExcelFile Name:'};
dlg_title = 'Input';
num_lines = 1;
answer = inputdlg(prompt,dlg_title,num_lines);
nameExcelFile = answer{1,1};

    if isempty(nameExcelFile)
        errordlg('Excel File not found','File Error');
        errorNameExcelFile = 'TRUE';
    else
        errorNameExcelFile = 'FALSE';
    end
end

if (strcmp(errorSubsystem,'FALSE') || strcmp(errorNameExcelFile,'FALSE'))    
	handles=find_system(nameSubsystem,'FindAll','On','SearchDepth',1,'BlockType','Inport'); 
	listInports=get(handles,'Name'); %list of inports

	handles=find_system(nameSubsystem,'FindAll','On','SearchDepth',1,'BlockType','Outport'); 
	listOutports=get(handles,'Name'); %list of outports

	listInOut={}; %list of inports and outports
	for ii=1:size(listInports,1)
		listInOut{ii,1} = listInports{ii,1};
	end

	for ii=1:size(listOutports,1)
		listInOut{ii,2} = listOutports{ii,1};
	end
	
	%Summary Excel File
	xlsFileResult_synthese = fullfile(pathRep,[nameExcelFile '.xls']);
	
	fidOutput_synthese = fopen(xlsFileResult_synthese,'w');  

	%Excel file header
	fprintf(fidOutput_synthese,'%s\n', ['Subsystem Name: ' nameSubsystem]);
	fprintf(fidOutput_synthese,'%s\t','Inports');
	fprintf(fidOutput_synthese,'%s\n','Outports');

	%Filling the file		
	for tt = 1 : size(listInOut,1)
		fprintf(fidOutput_synthese,'%s',listInOut{tt,1});
        fprintf(fidOutput_synthese,'\t');
		fprintf(fidOutput_synthese,'%s',listInOut{tt,2});  
        fprintf(fidOutput_synthese,'\n');
	end
	
	messagebox = msgbox('Operation Completed!','Success');
	disp('Recovery of the parameters carried out successfully! Generated Summary Excel File:')
	disp('')
	disp(xlsFileResult_synthese)
	disp('')	  
		
	fclose(fidOutput_synthese); 
end