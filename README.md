# panel-projesi-kodlar-
kodların çalışması için command window kodları gerek

run_pm10_city_region_v4.m kodu için :

clear functions
baseFolder = "C:\Users\User\Desktop\proje\hava kirliliği verileri";
addpath(genpath(baseFolder));

[StationSeason, CitySeason, RegionSeason] = run_pm10_city_region_v4(baseFolder, 2025);


make_region_7x4_table.m için:

outputsDir = "C:\Users\User\Desktop\proje\hava kirliliği verileri\outputs";
make_region_7x4_table(outputsDir, 2025);
