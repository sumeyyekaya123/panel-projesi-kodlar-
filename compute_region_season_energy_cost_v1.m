function RegionSeasonEnergyCost = compute_region_season_energy_cost_v1(outputsFolder, targetYear)
% compute_region_season_energy_cost_v1
%
% CitySeasonEnergyCostLoss dosyasindan
% bolge + mevsim bazinda ortalama alir.
%
% Girdi:
%   outputsFolder = "C:\Users\User\Desktop\2. kısım\outputs";
%   targetYear    = 2025
%
% Beklenen dosya:
%   CitySeasonEnergyCostLoss_2025.xlsx
%
% Cikti:
%   RegionSeasonEnergyCostLoss_2025.xlsx

    arguments
        outputsFolder (1,1) string
        targetYear (1,1) double
    end

    % -----------------------------
    % Girdi dosyasi
    % -----------------------------
    inFile = fullfile(outputsFolder, sprintf("CitySeasonEnergyCostLoss_%d.xlsx", targetYear));

    if ~exist(inFile, "file")
        error("Girdi dosyasi bulunamadi: %s", inFile);
    end

    % -----------------------------
    % Excel oku
    % -----------------------------
    T = readtable(inFile, "PreserveVariableNames", true);

    requiredVars = { ...
        'Region','City','Season','Year', ...
        'SL_s_pct','E_clean_s_kWh_per_kWp', ...
        'E_loss_s_kWh_per_kWp','Cost_loss_s_TL_per_kWp'};

    for i = 1:numel(requiredVars)
        if ~ismember(requiredVars{i}, T.Properties.VariableNames)
            error("Gerekli kolon eksik: %s", requiredVars{i});
        end
    end

    % -----------------------------
    % Guvenli veri donusumu
    % -----------------------------
    T.Region = lower(strtrim(string(T.Region)));
    T.City   = lower(strtrim(string(T.City)));
    T.Season = upper(strtrim(string(T.Season)));

    T.SL_s_pct = force_numeric(T.SL_s_pct);
    T.E_clean_s_kWh_per_kWp = force_numeric(T.E_clean_s_kWh_per_kWp);
    T.E_loss_s_kWh_per_kWp = force_numeric(T.E_loss_s_kWh_per_kWp);
    T.Cost_loss_s_TL_per_kWp = force_numeric(T.Cost_loss_s_TL_per_kWp);

    % Gecersiz satirlari atla
    good = isfinite(T.SL_s_pct) & ...
           isfinite(T.E_clean_s_kWh_per_kWp) & ...
           isfinite(T.E_loss_s_kWh_per_kWp) & ...
           isfinite(T.Cost_loss_s_TL_per_kWp);

    T = T(good, :);

    if isempty(T)
        error("Gecerli veri kalmadi. Sayisal kolonlar kontrol edilmeli.");
    end

    % Fiziksel sinirlar
    T.SL_s_pct(T.SL_s_pct < 0) = 0;
    T.SL_s_pct(T.SL_s_pct > 100) = 100;

    T.E_clean_s_kWh_per_kWp(T.E_clean_s_kWh_per_kWp < 0) = 0;
    T.E_loss_s_kWh_per_kWp(T.E_loss_s_kWh_per_kWp < 0) = 0;
    T.Cost_loss_s_TL_per_kWp(T.Cost_loss_s_TL_per_kWp < 0) = 0;

    % Sadece beklenen mevsimler kalsin
    validSeason = ismember(T.Season, ["DJF","MAM","JJA","SON"]);
    T = T(validSeason, :);

    if isempty(T)
        error("Gecerli mevsim etiketi kalmadi.");
    end

    % -----------------------------
    % Bolge + mevsim bazinda grupla
    % -----------------------------
    [G, regU, seasonU, yearU] = findgroups(T.Region, T.Season, T.Year);

    meanSL = splitapply(@mean, T.SL_s_pct, G);
    meanEClean = splitapply(@mean, T.E_clean_s_kWh_per_kWp, G);
    meanELoss = splitapply(@mean, T.E_loss_s_kWh_per_kWp, G);
    meanCostLoss = splitapply(@mean, T.Cost_loss_s_TL_per_kWp, G);
    nCities = splitapply(@(x) numel(unique(x)), T.City, G);

    RegionSeasonEnergyCost = table();
    RegionSeasonEnergyCost.Region = regU;
    RegionSeasonEnergyCost.Season = seasonU;
    RegionSeasonEnergyCost.Year = yearU;
    RegionSeasonEnergyCost.mean_SL_s_pct = meanSL;
    RegionSeasonEnergyCost.mean_E_clean_s_kWh_per_kWp = meanEClean;
    RegionSeasonEnergyCost.mean_E_loss_s_kWh_per_kWp = meanELoss;
    RegionSeasonEnergyCost.mean_Cost_loss_s_TL_per_kWp = meanCostLoss;
    RegionSeasonEnergyCost.N_cities = nCities;

    RegionSeasonEnergyCost = sortrows(RegionSeasonEnergyCost, {'Region','Season'});

    % -----------------------------
    % Excel yaz
    % -----------------------------
    outFile = fullfile(outputsFolder, sprintf("RegionSeasonEnergyCostLoss_%d.xlsx", targetYear));
    writetable(RegionSeasonEnergyCost, outFile, "FileType", "spreadsheet");

    fprintf("Kaydedildi: %s\n", outFile);
end

function y = force_numeric(x)
    if isnumeric(x)
        y = double(x);
        return;
    end

    if iscell(x)
        x = string(x);
    end

    if isstring(x) || ischar(x)
        x = strrep(string(x), ",", ".");
        y = str2double(x);
        return;
    end

    y = nan(size(x));
end