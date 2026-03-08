function CitySeasonalSL = compute_city_seasonal_sl_v1(outputsFolder, targetYear)
% compute_city_seasonal_sl_v1
%
% 9) Mevsimsel soiling yuzdesi
%
% Formul:
%   SL_s(%) = (1 / N_d,s) * sum(SL_d(%))
%
% Girdi:
%   outputsFolder = "C:\Users\User\Desktop\2. kısım\outputs";
%   targetYear    = 2025
%
% Beklenen dosya:
%   CityDailySR_2025.xlsx
%
% Cikti:
%   CitySeasonalSL_2025.xlsx

    arguments
        outputsFolder (1,1) string
        targetYear (1,1) double
    end

    % -----------------------------
    % Girdi dosyasi
    % -----------------------------
    inFile = fullfile(outputsFolder, sprintf("CityDailySR_%d.xlsx", targetYear));

    if ~exist(inFile, "file")
        error("Girdi dosyasi bulunamadi: %s", inFile);
    end

    % -----------------------------
    % Excel oku
    % -----------------------------
    T = readtable(inFile, "PreserveVariableNames", true);

    requiredVars = {'Region','City','Date','Year','SL_t_pct'};
    for i = 1:numel(requiredVars)
        if ~ismember(requiredVars{i}, T.Properties.VariableNames)
            error("Gerekli kolon eksik: %s", requiredVars{i});
        end
    end

    % Date ve SL_t_pct guvenli okuma
    T.Date = force_datetime(T.Date);
    T.SL_t_pct = force_numeric(T.SL_t_pct);

    % Okunamayan / gecersiz satirlari atla
    good = ~isnat(T.Date) & isfinite(T.SL_t_pct);
    T = T(good, :);

    if isempty(T)
        error("Gecerli veri kalmadi. Date veya SL_t_pct kolonlari kontrol edilmeli.");
    end

    % Fiziksel sinirlar
    T.SL_t_pct(T.SL_t_pct < 0) = 0;
    T.SL_t_pct(T.SL_t_pct > 100) = 100;

    % -----------------------------
    % Mevsim ata
    % -----------------------------
    T.Season = strings(height(T), 1);

    m = month(T.Date);

    T.Season(ismember(m, [12 1 2]))    = "DJF";
    T.Season(ismember(m, [3 4 5]))     = "MAM";
    T.Season(ismember(m, [6 7 8]))     = "JJA";
    T.Season(ismember(m, [9 10 11]))   = "SON";

    % -----------------------------
    % Sehir + mevsim bazinda ortalama
    % -----------------------------
    [G, regU, cityU, seasonU, yearU] = findgroups( ...
        T.Region, T.City, T.Season, T.Year);

    meanSL = splitapply(@mean, T.SL_t_pct, G);
    nDays  = splitapply(@numel, T.SL_t_pct, G);

    CitySeasonalSL = table();
    CitySeasonalSL.Region = regU;
    CitySeasonalSL.City = cityU;
    CitySeasonalSL.Season = seasonU;
    CitySeasonalSL.Year = yearU;
    CitySeasonalSL.SL_s_pct = meanSL;
    CitySeasonalSL.N_d_s = nDays;

    CitySeasonalSL = sortrows(CitySeasonalSL, {'Region','City','Season'});

    % -----------------------------
    % Excel yaz
    % -----------------------------
    outFile = fullfile(outputsFolder, sprintf("CitySeasonalSL_%d.xlsx", targetYear));
    writetable(CitySeasonalSL, outFile, "FileType", "spreadsheet");

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