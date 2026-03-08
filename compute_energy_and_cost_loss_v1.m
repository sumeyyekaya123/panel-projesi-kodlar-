function CitySeasonEnergyCost = compute_energy_and_cost_loss_v1(outputsFolder, targetYear)
% compute_energy_and_cost_loss_v1
%
% 10) Soiling nedeniyle enerji kaybi
% 11) Soiling nedeniyle parasal kayip
%
% Kullanilan varsayimlar:
%   E_clean_year = 1550 kWh/kWp-year
%   p_e = 2.5904 TL/kWh
%
% Girdi:
%   outputsFolder = "C:\Users\User\Desktop\2. kısım\outputs";
%   targetYear    = 2025
%
% Beklenen dosya:
%   CitySeasonalSL_2025.xlsx
%
% Cikti:
%   CitySeasonEnergyCostLoss_2025.xlsx
%
% Not:
% Bu surum yillik toplam temiz uretimi mevsimlere gun sayisina gore dagitir.

    arguments
        outputsFolder (1,1) string
        targetYear (1,1) double
    end

    % -----------------------------
    % Varsayimlar
    % -----------------------------
    E_clean_year_kWh_per_kWp = 1550;
    p_e_TL_per_kWh = 2.5904;

    % -----------------------------
    % Girdi dosyasi
    % -----------------------------
    inFile = fullfile(outputsFolder, sprintf("CitySeasonalSL_%d.xlsx", targetYear));

    if ~exist(inFile, "file")
        error("Girdi dosyasi bulunamadi: %s", inFile);
    end

    % -----------------------------
    % Excel oku
    % -----------------------------
    T = readtable(inFile, "PreserveVariableNames", true);

    requiredVars = {'Region','City','Season','Year','SL_s_pct','N_d_s'};
    for i = 1:numel(requiredVars)
        if ~ismember(requiredVars{i}, T.Properties.VariableNames)
            error("Gerekli kolon eksik: %s", requiredVars{i});
        end
    end

    % -----------------------------
    % Guvenli veri donusumu
    % -----------------------------
    T.Season = upper(strtrim(string(T.Season)));
    T.SL_s_pct = force_numeric(T.SL_s_pct);
    T.N_d_s = force_numeric(T.N_d_s);

    % Okunamayan/gecersiz satirlari atla
    good = isfinite(T.SL_s_pct) & isfinite(T.N_d_s) & (T.N_d_s > 0);
    T = T(good, :);

    if isempty(T)
        error("Gecerli veri kalmadi. SL_s_pct veya N_d_s kolonlari kontrol edilmeli.");
    end

    % Fiziksel sinirlar
    T.SL_s_pct(T.SL_s_pct < 0) = 0;
    T.SL_s_pct(T.SL_s_pct > 100) = 100;

    % -----------------------------
    % Mevsim gun sayilari
    % -----------------------------
    seasonDays = zeros(height(T), 1);

    for i = 1:height(T)
        s = T.Season(i);

        switch s
            case "DJF"
                seasonDays(i) = 90;
            case "MAM"
                seasonDays(i) = 92;
            case "JJA"
                seasonDays(i) = 92;
            case "SON"
                seasonDays(i) = 91;
            otherwise
                error("Bilinmeyen mevsim etiketi: %s", s);
        end
    end

    % -----------------------------
    % Kirlenmesiz mevsimsel enerji
    % -----------------------------
    T.E_clean_s_kWh_per_kWp = E_clean_year_kWh_per_kWp .* (seasonDays / 365);

    % -----------------------------
    % Mevsimsel kaybi kesre cevir
    % -----------------------------
    T.SL_s_frac = T.SL_s_pct ./ 100;

    % -----------------------------
    % Enerji kaybi
    % -----------------------------
    T.E_loss_s_kWh_per_kWp = T.SL_s_frac .* T.E_clean_s_kWh_per_kWp;

    % -----------------------------
    % Parasal kayip
    % -----------------------------
    T.p_e_TL_per_kWh = repmat(p_e_TL_per_kWh, height(T), 1);
    T.Cost_loss_s_TL_per_kWp = T.E_loss_s_kWh_per_kWp .* T.p_e_TL_per_kWh;

    % -----------------------------
    % Son tablo
    % -----------------------------
    CitySeasonEnergyCost = table();
    CitySeasonEnergyCost.Region = T.Region;
    CitySeasonEnergyCost.City = T.City;
    CitySeasonEnergyCost.Season = T.Season;
    CitySeasonEnergyCost.Year = T.Year;
    CitySeasonEnergyCost.SL_s_pct = T.SL_s_pct;
    CitySeasonEnergyCost.SL_s_frac = T.SL_s_frac;
    CitySeasonEnergyCost.E_clean_s_kWh_per_kWp = T.E_clean_s_kWh_per_kWp;
    CitySeasonEnergyCost.E_loss_s_kWh_per_kWp = T.E_loss_s_kWh_per_kWp;
    CitySeasonEnergyCost.p_e_TL_per_kWh = T.p_e_TL_per_kWh;
    CitySeasonEnergyCost.Cost_loss_s_TL_per_kWp = T.Cost_loss_s_TL_per_kWp;
    CitySeasonEnergyCost.N_d_s = T.N_d_s;

    CitySeasonEnergyCost = sortrows(CitySeasonEnergyCost, {'Region','City','Season'});

    % -----------------------------
    % Excel yaz
    % -----------------------------
    outFile = fullfile(outputsFolder, sprintf("CitySeasonEnergyCostLoss_%d.xlsx", targetYear));
    writetable(CitySeasonEnergyCost, outFile, "FileType", "spreadsheet");

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