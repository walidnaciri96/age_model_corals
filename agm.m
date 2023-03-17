function [te_m,year_month] = agm(myrange1,myrange2,path)
%   myrange1 is the range of the TE data in the excel file (e.g.
%   'C1:F266')
%   myrange2 is the range of parameters in the excel file (e.g. 'A1:E417')

%   "path" is the entire path to the excel file 
%    
%%%%%%%%    File format: 

%   - data should be in a sheet called "original_res"
%   - an empty sheet for monthly res data this script will create 
%   - a parameters sheet called 'age model' with Depth, Date_array (an array with the entire monthly resolved record,
%   AP_array (an array with all anchor points), AP_depth_array (an array with all anchor points' depths)
    


data = readtable(path, 'Sheet', 'original res', 'Range', myrange1, 'PreserveVariableNames', true);
params = readtable(path, 'Sheet', 'age model', 'Range', myrange2, 'PreserveVariableNames', true);


te = [];
for n = 1:width(data)
    x = data(:,n);
    te = [te,x];    %Matrix of TE values
end
clear x 

marker_depth = rmmissing(params.AP_depth_array); %A matrix of the depth of each anchor point 
marker_year = rmmissing(params.AP_array); %A matrix of the year of each anchor point 
coral_depth = rmmissing(params.Depth); %A matrix of the entire coral depth 
year_month = rmmissing(params.Date_array); %A matrix of time (organised monthly) form the begining to the end of the core

for n = 1:(length(marker_year)-1)
    if marker_year(n)>marker_year(n+1) ~= 1
    disp('Error with the anchor points, check that the vector is striclty increasing around this value:')
    disp(n+1)
    plot(marker_year)
    return
    else
    end
end

%Interpolation
year=interp1(marker_depth,marker_year,coral_depth,'linear','extrap'); %Interpolates the time array of the original data
%writematrix(year,path,'Sheet','original_res','Range','B2:B3028','AutoFitWidth',0);

%Interpolation to monthly values
te_m=[];
for n=1:width(data) 
    x=table2array(te(:,n));
    xi=interp1(year,x,year_month,'linear','extrap');
    te_m=[te_m,xi]; % Monthly based TE matrix

end

writematrix(te_m,path,'Sheet','m res','Range','C2:K3028','AutoFitWidth',0);

writematrix(year_month,path,'Sheet','m res','Range','A2:A3028','AutoFitWidth',0);
end

