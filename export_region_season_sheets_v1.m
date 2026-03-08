function outFile = export_region_season_sheets_v1(outputsFolder, targetYear)
% RegionSeasonEnergyCostLoss dosyasini okuyup
% her mevsimi ayri Excel sayfasina yazar.
%
% Sheet adlari:
%   Kış, İlkbahar, Yaz, Sonbahar
%
% Dosya adi degismez:
%   RegionSeasonEnergyCostLoss_BySeasonSheets_2025.xlsx

    arguments
        outputsFolder (1,1) string
        targetYear (1,1) double
    end

    % Girdi dosyasi
    inFile = fullfile(outputsFolder, sprintf("RegionSeasonEnergyCostLoss_%d.xlsx", targetYear));

    if ~exist(inFile, "file")
        error("Girdi dosyasi bulunamadi: %s", inFile);
    end

    % Excel oku
    T = readtable(inFile, "PreserveVariableNames", true);

    requiredVars = { ...
        'Region','Season','Year', ...
        'mean_SL_s_pct', ...
        'mean_E_clean_s_kWh_per_kWp', ...
        'mean_E_loss_s_kWh_per_kWp', ...
        'mean_Cost_loss_s_TL_per_kWp', ...
        'N_cities'};

    for i = 1:numel(requiredVars)
        if ~ismember(requiredVars{i}, T.Properties.VariableNames)
            error("Gerekli kolon eksik: %s", requiredVars{i});
        end
    end

    % Guvenli donusum
    T.Region = lower(strtrim(string(T.Region)));
    T.Season = upper(strtrim(string(T.Season)));

    T.mean_SL_s_pct = force_numeric(T.mean_SL_s_pct);
    T.mean_E_clean_s_kWh_per_kWp = force_numeric(T.mean_E_clean_s_kWh_per_kWp);
    T.mean_E_loss_s_kWh_per_kWp = force_numeric(T.mean_E_loss_s_kWh_per_kWp);
    T.mean_Cost_loss_s_TL_per_kWp = force_numeric(T.mean_Cost_loss_s_TL_per_kWp);
    T.N_cities = force_numeric(T.N_cities);

    % Gecersiz satirlari at
    good = isfinite(T.mean_SL_s_pct) & ...
           isfinite(T.mean_E_clean_s_kWh_per_kWp) & ...
           isfinite(T.mean_E_loss_s_kWh_per_kWp) & ...
           isfinite(T.mean_Cost_loss_s_TL_per_kWp) & ...
           isfinite(T.N_cities);

    T = T(good, :);

    if isempty(T)
        error("Gecerli veri kalmadi.");
    end

    % Sadece beklenen mevsimler
    validSeasons = ["DJF","MAM","JJA","SON"];
    T = T(ismember(T.Season, validSeasons), :);

    if isempty(T)
        error("Beklenen mevsim verisi bulunamadi.");
    end

    % Ayni bolge+mevsim birden fazla ise tek satira indir
    [G, regU, seasonU, yearU] = findgroups(T.Region, T.Season, T.Year);

    TT = table();
    TT.Region = regU;
    TT.Season = seasonU;
    TT.Year = yearU;
    TT.mean_SL_s_pct = splitapply(@mean, T.mean_SL_s_pct, G);
    TT.mean_E_clean_s_kWh_per_kWp = splitapply(@mean, T.mean_E_clean_s_kWh_per_kWp, G);
    TT.mean_E_loss_s_kWh_per_kWp = splitapply(@mean, T.mean_E_loss_s_kWh_per_kWp, G);
    TT.mean_Cost_loss_s_TL_per_kWp = splitapply(@mean, T.mean_Cost_loss_s_TL_per_kWp, G);
    TT.N_cities = splitapply(@max, T.N_cities, G);

    % Season kolonunu Turkce yap
    TT.Season(TRowMatch(TT.Season, "DJF")) = "Kış";
    TT.Season(TRowMatch(TT.Season, "MAM")) = "İlkbahar";
    TT.Season(TRowMatch(TT.Season, "JJA")) = "Yaz";
    TT.Season(TRowMatch(TT.Season, "SON")) = "Sonbahar";

    % Cikti dosyasi
    outFile = fullfile(outputsFolder, sprintf("RegionSeasonEnergyCostLoss_BySeasonSheets_%d.xlsx", targetYear));

    if exist(outFile, "file")
        delete(outFile);
    end

    % Sheetlere ayir
    S1 = TT(TT.Season == "Kış", :);
    S2 = TT(TT.Season == "İlkbahar", :);
    S3 = TT(TT.Season == "Yaz", :);
    S4 = TT(TT.Season == "Sonbahar", :);

    S1 = sortrows(S1, {'Region'});
    S2 = sortrows(S2, {'Region'});
    S3 = sortrows(S3, {'Region'});
    S4 = sortrows(S4, {'Region'});

    writetable(S1, outFile, "FileType", "spreadsheet", "Sheet", "Kış");
    writetable(S2, outFile, "FileType", "spreadsheet", "Sheet", "İlkbahar");
    writetable(S3, outFile, "FileType", "spreadsheet", "Sheet", "Yaz");
    writetable(S4, outFile, "FileType", "spreadsheet", "Sheet", "Sonbahar");

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

function idx = TRowMatch(x, val)
    idx = (string(x) == string(val));
end