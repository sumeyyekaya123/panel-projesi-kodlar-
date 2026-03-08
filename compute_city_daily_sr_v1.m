function CityDailySR = compute_city_daily_sr_v1(outputsFolder, targetYear)
% compute_city_daily_sr_v1
%
% 7) Birikimden gunluk soiling kaybi
% 8) Soiling ratio
%
% Kullanilan formul:
%   SL_t(%) = 34.37 * erf(0.17 * omega_t^0.8473)
%   SR_t    = 1 - SL_t(%)/100
%
% Girdi:
%   outputsFolder = "C:\Users\User\Desktop\2. kısım\outputs";
%   targetYear    = 2025
%
% Beklenen dosya:
%   CityDailyOmegaReset_2025.xlsx
%
% Cikti:
%   CityDailySR_2025.xlsx

    arguments
        outputsFolder (1,1) string
        targetYear (1,1) double
    end

    % -----------------------------
    % Girdi dosyasi
    % -----------------------------
    inFile = fullfile(outputsFolder, sprintf("CityDailyOmegaReset_%d.xlsx", targetYear));

    if ~exist(inFile, "file")
        error("Girdi dosyasi bulunamadi: %s", inFile);
    end

    % -----------------------------
    % Excel oku
    % -----------------------------
    T = readtable(inFile, "PreserveVariableNames", true);

    requiredVars = {'Region','City','Date','Year'};
    for i = 1:numel(requiredVars)
        if ~ismember(requiredVars{i}, T.Properties.VariableNames)
            error("Gerekli kolon eksik: %s", requiredVars{i});
        end
    end

    % Oncelik reset uygulanmis omega
    if ismember('omega_t_reset_gm2', T.Properties.VariableNames)
        omegaVals = T.omega_t_reset_gm2;
        omegaVarName = 'omega_t_reset_gm2';
    elseif ismember('omega_t_gm2', T.Properties.VariableNames)
        omegaVals = T.omega_t_gm2;
        omegaVarName = 'omega_t_gm2';
    else
        error("Omega kolonu bulunamadi. 'omega_t_reset_gm2' veya 'omega_t_gm2' gerekli.");
    end

    % Date tipi
    T.Date = force_datetime(T.Date);

    % Omega sayisala zorla
    omegaVals = force_numeric(omegaVals);

    % Okunamayan / gecersiz satirlari atla
    good = ~isnat(T.Date) & isfinite(omegaVals);
    T = T(good, :);
    omegaVals = omegaVals(good);

    if isempty(T)
        error("Gecerli veri kalmadi. Date veya omega kolonlari kontrol edilmeli.");
    end

    % Fiziksel guvenlik: negatif omega varsa sifira cek
    omegaVals(omegaVals < 0) = 0;

    % -----------------------------
    % Formuller
    % -----------------------------
    SL_t_pct = 34.37 .* erf(0.17 .* (omegaVals .^ 0.8473));
    SR_t = 1 - (SL_t_pct ./ 100);

    % Fiziksel sinirlar
    SL_t_pct(SL_t_pct < 0) = 0;
    SL_t_pct(SL_t_pct > 100) = 100;

    SR_t(SR_t < 0) = 0;
    SR_t(SR_t > 1) = 1;

    % -----------------------------
    % Cikti tablosu
    % -----------------------------
    CityDailySR = table();
    CityDailySR.Region = T.Region;
    CityDailySR.City = T.City;
    CityDailySR.Date = T.Date;
    CityDailySR.Year = T.Year;

    CityDailySR.(omegaVarName) = omegaVals;
    CityDailySR.SL_t_pct = SL_t_pct;
    CityDailySR.SR_t = SR_t;
    CityDailySR.OneMinusSR_t = 1 - CityDailySR.SR_t;

    CityDailySR = sortrows(CityDailySR, {'Region','City','Date'});

    % -----------------------------
    % Excel yaz
    % -----------------------------
    outFile = fullfile(outputsFolder, sprintf("CityDailySR_%d.xlsx", targetYear));
    writetable(CityDailySR, outFile, "FileType", "spreadsheet");

    fprintf("Kaydedildi: %s\n", outFile);
end

% ==========================================================
% Yardimci fonksiyonlar
% ==========================================================

function dt = force_datetime(x)
    if isdatetime(x)
        dt = x;
        return;
    end

    if iscell(x)
        x = string(x);
    end

    if isstring(x) || ischar(x)
        try
            dt = datetime(x, "InputFormat", "yyyy-MM-dd");
            return;
        catch
        end
        try
            dt = datetime(x, "InputFormat", "yyyy-MM-dd HH:mm:ss");
            return;
        catch
        end
        try
            dt = datetime(x, "InputFormat", "dd.MM.yyyy");
            return;
        catch
        end
        try
            dt = datetime(x, "InputFormat", "dd.MM.yyyy HH:mm:ss");
            return;
        catch
        end
        try
            dt = datetime(x, "InputFormat", "dd.MM.yyyy HH:mm");
            return;
        catch
        end
        try
            dt = datetime(x);
            return;
        catch
        end
    end

    if isnumeric(x)
        try
            dt = datetime(x, "ConvertFrom", "excel");
            return;
        catch
        end
    end

    dt = NaT(size(x));
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