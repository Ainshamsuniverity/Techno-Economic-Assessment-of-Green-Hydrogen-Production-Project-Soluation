tic
clc; %clear command window
clear; %clear workspace
load('initData.mat') %Load the initial data

%%  Specify the Excel file name

irradianceFile = 'irradiance.xlsx'; %get the irradiance file
tempratureFile = 'temprature.xlsx'; %get the temperature file
windSpeedFile = 'wind speed.xlsx'; %get the wind speed file
electricityPricesFile='electricityPrices.xlsx'; %get the electricity prices grid file
countryFile= 'Country.xlsx'; %get the countries name file

%Import all sheets data from Excel
irradiance = readtable(irradianceFile);
electricityPrices=readtable(electricityPricesFile);
temprature = readtable(tempratureFile);
windSpeed = readtable(windSpeedFile);
country = readtable(countryFile);
CountryNameData = country.Countries; % get the first column from the file
OurStationData = struct('Country_Name', CountryNameData); % create the station data struct data

% get the structure data in
 for k=1:242
OurStationData(k).ElectricityPrices = electricityPrices.(k);
OurStationData(k).Irradiance = irradiance.(k);
OurStationData(k).Temprature = temprature.(k);
OurStationData(k).WindSpeed = windSpeed.(k);
 end

%% Running the code many times and get Egrid out

irradiance=zeros(8760,1); % make the irradiance matrix zeros
 for i=1:242 % For loop to repeat the step
irradiance=((OurStationData(i).Irradiance*480)+1000); % copy the data from station data to the irradiance  
model = 'Green_Hydrogen_Model'; % call the model
open_system(model); % Load the Simscape model
sim('Green_Hydrogen_Model'); % run the simulation with the copied data
Egrid_copy(i)=ans.Egrid; % save the grid energy in Egrid_copy
 end

%% Cost Analysis

for i=1:242

Egrid=Egrid_copy(i).signals.values(8760,1);
cost_grid(i) = Egrid * OurStationData(i).ElectricityPrices;
cost_grid_in_coloumns = repmat(cost_grid(1, :), 8760, 1); % make the grid cost in the whole column
Total_Cost_of_grid=ans.Total_Cost_without_Grid.signals.values - cost_grid; % Total cost calculation
Total_cost(i)=(sum(Total_Cost_of_grid(:,i))/8760); % averaging of total cost
H2_cost_measure = max(Total_cost); % measure max cost
if H2_cost_measure < 0 % if loop to get the max and min cost with the suitable sign
    min_cost_H2 = H2_cost_measure;
    max_cost_H2=min(Total_cost);
else
    max_cost_H2=H2_cost_measure;
    min_cost_H2=min(Total_cost);
end
end

fprintf('Max Cost: %f\n', max_cost_H2); % print the max cost
[~, colNoMaxH2] = find(Total_cost == max_cost_H2); % get the index for the max cost
disp('Country_Name'); % print the country name
disp(OurStationData(colNoMaxH2).Country_Name); % print the country name
fprintf('Min Cost: %f\n', min_cost_H2); % print the min cost
[~, colNoMinH2] = find(Total_cost == min_cost_H2); % get the index for the min cost
disp('Country_Name'); % print the country name
disp(OurStationData(colNoMinH2).Country_Name); % print the country name

%% H2 Cost

Total_H2_in_year= ans.H2_Kg.signals.values(8760,1);
H2_Kg_cost= (Total_cost)/Total_H2_in_year;

H2_cost_measure = max(H2_Kg_cost); % measure max cost
if H2_cost_measure < 0 % if loop to get the max and min cost with the suitable sign
    min_cost_H2 = H2_cost_measure;
    max_cost_H2=min(H2_Kg_cost);
else
    max_cost_H2=H2_cost_measure;
    min_cost_H2=min(H2_Kg_cost);
end


fprintf('Max Cost H2: %f\n', max_cost_H2); % print the max cost
[rowNoMaxH2, colNoMaxH2] = find(H2_Kg_cost == max_cost_H2); % get the index for the max cost
disp('Country_Name'); % print the country name
disp(OurStationData(colNoMaxH2).Country_Name); % print the country name
fprintf('Min Cost H2: %f\n', min_cost_H2); % print the min cost
[rowNoMinH2, colNoMinH2] = find(H2_Kg_cost == min_cost_H2); % get the index for the min cost
disp('Country_Name'); % print the country name
disp(OurStationData(colNoMinH2).Country_Name); % print the country name
toc;
