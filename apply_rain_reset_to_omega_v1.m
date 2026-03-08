function [CityDailyRainMerged, CityDailyOmegaReset] = apply_rain_reset_to_omega_v1(outputsFolder, rainFolder, targetYear)
% apply_rain_reset_to_omega_v1
%
% 6) Dogal temizlik (yagisla reset)
% Kural:
%   If R_d > R_th, then omega_t = 0

    arguments
        outputsFolder (1,1) string
        rainFolder (1,1) string
        targetYear (1,1) double
    end

    % Parametre
    R_th_mm = 1;

    % -------------------------------------------------
    % Omega dosyasi
    % -------------------------------------------------
    omegaFile = fullfile(outputsFolder, sprintf("CityDailyOmega_%d.xlsx", targetYear));
    if ~exist(omegaFile, "file")
        error("Girdi dosyasi bulunamadi: %s", omegaFile);
    end

    Omega = readtable(omegaFile, "PreserveVariableNames", true);

    requiredVars = {'Region','City','Date','Year','omega_t_gm2','m_t_gm2'};
    for i = 1:numel(requiredVars)
        if ~ismember(requiredVars{i}, Omega.Properties.VariableNames)
            error("Omega dosyasinda gerekli kolon eksik: %s", requiredVars{i});
        end
    end

    if ~isdatetime(Omega.Date)
        Omega.Date = force_datetime(Omega.Date);
    end

    Omega.Region = lower(strtrim(string(Omega.Region)));
    Omega.City   = normalize_city_name(Omega.City);
    Omega = sortrows(Omega, {'Region','City','Date'});

    % -------------------------------------------------
    % Yagis klasorundeki bolge alt klasorlerini tara
    % -------------------------------------------------
    d = dir(rainFolder);
    d = d([d.isdir]);
    regionNames = string({d.name});
    regionNames = regionNames(~ismember(regionNames, [".", ".."]));

    RainAll = table();

    for r = 1:numel(regionNames)
        region = lower(strtrim(string(regionNames(r))));
        regionPath = fullfile(rainFolder, regionNames(r));

        files = dir(fullfile(regionPath, "*.xlsx"));
        if isempty(files)
            warning("Bu bolge klasorunde xlsx yok: %s", regionPath);
            continue;
        end

        for k = 1:numel(files)
            filePath = fullfile(regionPath, files(k).name);

            try
                [~, cityName, ~] = fileparts(files(k).name);
                cityName = normalize_city_name(cityName);

                raw = readtable(filePath, "PreserveVariableNames", true);
                vnames = lower(strtrim(string(raw.Properties.VariableNames)));

                % date kolonu
                dateIdx = find(vnames == "date", 1);
                if isempty(dateIdx)
                    warning("date kolonu bulunamadi, dosya atlandi: %s", filePath);
                    continue;
                end

                % yagis kolonu: prcp veya prc
                prcpIdx = find(vnames == "prcp" | vnames == "prc", 1);
                if isempty(prcpIdx)
                    warning("prcp/prc kolonu bulunamadi, dosya atlandi: %s", filePath);
                    continue;
                end

                dt = raw{:, dateIdx};
                rain = raw{:, prcpIdx};

                dt = force_datetime(dt);
                rain = force_numeric(rain);

                % okunamayan / gecersiz / negatif degerleri atla
                good = ~isnat(dt) & isfinite(rain) & (rain >= 0);
                dt = dt(good);
                rain = rain(good);

                if isempty(dt)
                    warning("Yagis dosyasi atlandi (gecerli veri yok): %s", filePath);
                    continue;
                end

                keep = year(dt) == targetYear;
                dt = dt(keep);
                rain = rain(keep);

                if isempty(dt)
                    continue;
                end

                dt = dateshift(dt(:), "start", "day");

                T = table();
                T.Region = repmat(region, numel(dt), 1);
                T.City   = repmat(cityName, numel(dt), 1);
                T.Date   = dt;
                T.Year   = repmat(targetYear, numel(dt), 1);
                T.R_d_mm = rain(:);

                % Ayni gun birden fazla satir varsa gunluk toplam
                [G, regU, cityU, dateU, yearU] = findgroups(T.Region, T.City, T.Date, T.Year);
                dailyRain = splitapply(@sum, T.R_d_mm, G);

                CityRain = table();
                CityRain.Region = regU;
                CityRain.City   = cityU;
                CityRain.Date   = dateU;
                CityRain.Year   = yearU;
                CityRain.R_d_mm = dailyRain;

                RainAll = [RainAll; CityRain]; %#ok<AGROW>

            catch ME
                warning("Yagis dosyasi islenemedi: %s\nSebep: %s", filePath, ME.message);
            end
        end
    end

    if isempty(RainAll)
        error("Hic yagis verisi okunamadi.");
    end

    % -------------------------------------------------
    % Tekrarlari birlestir
    % -------------------------------------------------
    [G2, regU2, cityU2, dateU2, yearU2] = findgroups(RainAll.Region, RainAll.City, RainAll.Date, RainAll.Year);
    rainSum = splitapply(@sum, RainAll.R_d_mm, G2);

    RainDaily = table();
    RainDaily.Region = regU2;
    RainDaily.City   = cityU2;
    RainDaily.Date   = dateU2;
    RainDaily.Year   = yearU2;
    RainDaily.R_d_mm = rainSum;

    RainDaily = sortrows(RainDaily, {'Region','City','Date'});

    % -------------------------------------------------
    % Omega ile yagisi birlestir
    % -------------------------------------------------
    CityDailyRainMerged = outerjoin( ...
        Omega, RainDaily, ...
        'Keys', {'Region','City','Date','Year'}, ...
        'MergeKeys', true, ...
        'Type', 'left');

    if ~ismember('R_d_mm', CityDailyRainMerged.Properties.VariableNames)
        error("Birlestirme sonrasi R_d_mm kolonu olusmadi.");
    end

    % Eksik yagis = 0
    CityDailyRainMerged.R_d_mm(isnan(CityDailyRainMerged.R_d_mm)) = 0;
    CityDailyRainMerged.R_th_mm = repmat(R_th_mm, height(CityDailyRainMerged), 1);
    CityDailyRainMerged.ResetFlag = CityDailyRainMerged.R_d_mm > CityDailyRainMerged.R_th_mm;

    CityDailyRainMerged = sortrows(CityDailyRainMerged, {'Region','City','Date'});

    % -------------------------------------------------
    % Reset uygulanmis omega
    % -------------------------------------------------
    CityDailyOmegaReset = CityDailyRainMerged;
    CityDailyOmegaReset.omega_t_reset_gm2 = zeros(height(CityDailyOmegaReset), 1);

    [G3, ~, ~] = findgroups(CityDailyOmegaReset.Region, CityDailyOmegaReset.City);

    for g = 1:max(G3)
        idx = find(G3 == g);
        [~, ord] = sort(CityDailyOmegaReset.Date(idx));
        idx = idx(ord);

        omegaVals = zeros(numel(idx), 1);

        for j = 1:numel(idx)
            thisIdx = idx(j);

            if j == 1
                prevOmega = 0;
            else
                prevOmega = omegaVals(j-1);
            end

            mt = CityDailyOmegaReset.m_t_gm2(thisIdx);
            resetFlag = CityDailyOmegaReset.ResetFlag(thisIdx);

            if resetFlag
                omegaVals(j) = 0;
            else
                omegaVals(j) = prevOmega + mt;
            end
        end

        CityDailyOmegaReset.omega_t_reset_gm2(idx) = omegaVals;
    end

    % -------------------------------------------------
    % Kolon duzeni
    % -------------------------------------------------
    preferredVars = { ...
        'Region','City','Date','Year', ...
        'C_t_gm3','m_t_gm2', ...
        'R_d_mm','R_th_mm','ResetFlag', ...
        'omega_t_gm2','omega_t_reset_gm2'};

    keepVars = preferredVars(ismember(preferredVars, CityDailyOmegaReset.Properties.VariableNames));
    CityDailyOmegaReset = CityDailyOmegaReset(:, keepVars);

    % -------------------------------------------------
    % Excel yaz
    % -------------------------------------------------
    out1 = fullfile(outputsFolder, sprintf("CityDailyRainMerged_%d.xlsx", targetYear));
    out2 = fullfile(outputsFolder, sprintf("CityDailyOmegaReset_%d.xlsx", targetYear));

    writetable(CityDailyRainMerged, out1, "FileType", "spreadsheet");
    writetable(CityDailyOmegaReset, out2, "FileType", "spreadsheet");

    fprintf("Kaydedildi: %s\n", out1);
    fprintf("Kaydedildi: %s\n", out2);
end

% ==========================================================
% Yardimci fonksiyonlar
% ==========================================================

function s = normalize_city_name(s)
    s = lower(string(s));
    s = strtrim(s);

    s = replace(s, "ç", "c");
    s = replace(s, "ğ", "g");
    s = replace(s, "ı", "i");
    s = replace(s, "i̇", "i");
    s = replace(s, "ö", "o");
    s = replace(s, "ş", "s");
    s = replace(s, "ü", "u");

    s = replace(s, "_", " ");
    s = replace(s, "-", " ");

    while any(contains(s, "  "))
        s = replace(s, "  ", " ");
    end
end

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