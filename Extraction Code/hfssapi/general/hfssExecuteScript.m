% ----------------------------------------------------------------------------
% function hfssExecuteScript(hfssExePath, ScriptFile, [iconOpt = true], 
%                            [runAndExit = true])
% 
% Description :
% -------------
% Runs Ansoft HFSS from within MATLAB with the '/RunScriptAndExit' (and the 
% optional '/Iconic') options to execute a given VB Script file and then exit.
%
% Parameters :
% ------------
% hfssExePath - the complete HFSS Executable path (see note for more info).
% ScriptFile  - the complete path of the VB Script file to be executed by
%               HFSS.
% [iconOpt]   - (optional) if set to false, HFSS will run in the normal mode,
%               else it will run in the "iconized" mode.
% [runAndExit] - (optional) if set to true, HFSS will run the given script
%                and exit, else if set to false, HFSS will run and script
%                and wait for user input.
% 
% Note :
% ------
% 1. The hfssExePath must contain the PROPER and COMPLETE executable path of
%    the HFSS program. If you followed the usual installation instructions for
%    installing HFSS, then this will be:
%    'C:\"Program Files"\Ansoft\HFSS9\hfss.exe' 
%    (The quotes (" ... ") are required for directory names that have spaces 
%    in them.
% 2. Don't forget to specify the complete path for the script file either. 
%    Also, make sure that the script file is "closed" (using fclose()) before
%    calling this function for proper results.
% 3. Another important thing, make sure you run the script once with
%    runAndExit set to false, so that you are completely sure that the
%    script does not have any errors before running the [runAndExit =
%    false] mode. If there are any errors in the script HFSS will silently
%    exit and there will be no way to figure out if the script ran properly
%    or not.
%
% Example :
% ---------
% fid = fopen('C:\temp\myAntenna.vbs', 'wt');
% ... 
% % Close the script prior to calling ExecuteScript.
% fclose(fid);
% % Execute the script.
% hfssExecuteScript('C:\"Program Files"\Ansoft\HFSS9\hfss.exe',  ...
%                   'C:\temp\myAntenna.vbs', ...
%                    false);
%
% ----------------------------------------------------------------------------

% ----------------------------------------------------------------------------
% This file is part of HFSS-MATLAB-API.
%
% HFSS-MATLAB-API is free software; you can redistribute it and/or modify it 
% under the terms of the GNU General Public License as published by the Free 
% Software Foundation; either version 2 of the License, or (at your option) 
% any later version.
%
% HFSS-MATLAB-API is distributed in the hope that it will be useful, but 
% WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY 
% or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU General Public License 
% for more details.
%
% You should have received a copy of the GNU General Public License along with
% Foobar; if not, write to the Free Software Foundation, Inc., 59 Temple 
% Place, Suite 330, Boston, MA  02111-1307  USA
%
% Copyright 2004, Vijay Ramasami (rvc@ku.edu)
% ----------------------------------------------------------------------------

function hfssExecuteScript(hfssExePath, ScriptFile, iconMode, runAndExit)

% arguments processor.
if (nargin < 2)
	error('Insufficient number of arguments !');
elseif (nargin < 3)
	iconMode = [];
    runAndExit = [];
end;

% default arguments.
if isempty(iconMode)
	iconMode = true;
end;
if isempty(runAndExit)
    runAndExit = true;
end;

% Setup Iconic Mode.
if (iconMode == true)
	iconStr = '/Iconic';
else
	iconStr = '';
end;

% Setup Run and Exit Mode.
if (runAndExit)
    runStr = '/RunScriptAndExit';
else
    runStr = '/RunScript';
end;

% Create the Command Path.
cmdHFSS = [hfssExePath, ' ', iconStr, ' ', runStr, ' ' ,  ...
           ScriptFile];

% Execute the Command.
disp(sprintf('    Running HFSS using %s ...\n', cmdHFSS));
[Status, Result] = system(cmdHFSS);
if (Status ~= 0)
	warning('HFSS Execution returned an error status !');
end;
