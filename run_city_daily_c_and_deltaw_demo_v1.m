function [CityDailyC, CityDailyMt] = run_city_daily_c_and_deltaw_demo_v1(baseFolder, targetYear)
% DENEMELIK SURUM
% DeltaW yerine m_t kullanildi

    arguments
        baseFolder (1,1) string
        targetYear (1,1) double
    end

    % Sabitler
    v_eff_ms  = 0.004;
    t_s       = 86400;
    theta_deg = 30;
    cos_theta = cosd(theta_deg);

    outputsDir = fullfile(baseFolder, "outputs");
    if ~exist(outputsDir, "dir")
        mkdir(outputsDir);
    end

    % Bölge klasörleri
    d = dir(baseFolder);
    d = d([d.isdir]);
    regionNames = string({d.name});
    regionNames = regionNames(~ismember(regionNames, [".", "..", "outputs"]));

    allStationDaily = table();

    for r = 1:numel(regionNames)
        region = regionNames(r);
        regionPath = fullfile(baseFolder, region);

        files = dir(fullfile(regionPath, "*.xlsx"));
        if isempty(files)
            warning("Bu klasorde xlsx yok: %s", regionPath);
            continue;
        end

        for k = 1:numel(files)
            filePath = fullfile(regionPath, files(k).name);

            try
                [city, station] = parse_city_station(files(k).name);

                raw = readtable(filePath, "PreserveVariableNames", true);

                if width(raw) < 2
                    warning("Atlandi (yetersiz kolon): %s", filePath);
                    continue;
                end

                dt = raw{:,1};
                pm = raw{:,2};

                dt = force_datetime(dt);
                pm = force_numeric(pm);

                % Negatif PM10 degerlerini alma
                good = ~isnat(dt) & ~isnan(pm) & (pm >= 0);
                dt = dt(good);
                pm = pm(good);

                if isempty(dt)
                    warning("Atlandi (gecerli veri yok): %s", filePath);
                    continue;
                end

                keep = year(dt) == targetYear;
                dt = dt(keep);
                pm = pm(keep);

                if isempty(dt)
                    continue;
                end

                % Saatlik -> günlük
                T = table(dt(:), pm(:), 'VariableNames', {'DateTime','PM10_ugm3'});
                T.Date = dateshift(T.DateTime, "start", "day");

                [G, uniqueDates] = findgroups(T.Date);
                dailyMean = splitapply(@mean, T.PM10_ugm3, G);
                dailyNum  = splitapply(@numel, T.PM10_ugm3, G);

                stationDaily = table();
                stationDaily.Region = repmat(region, numel(uniqueDates), 1);
                stationDaily.City = repmat(city, numel(uniqueDates), 1);
                stationDaily.Station = repmat(station, numel(uniqueDates), 1);
                stationDaily.Date = uniqueDates;
                stationDaily.Year = repmat(targetYear, numel(uniqueDates), 1);
                stationDaily.PM10_d_ugm3 = dailyMean;
                stationDaily.C_d_gm3 = stationDaily.PM10_d_ugm3 * 1e-6;
                stationDaily.N_hours = dailyNum;

                allStationDaily = [allStationDaily; stationDaily]; %#ok<AGROW>

            catch ME
                warning("Dosya islenemedi: %s\nSebep: %s", filePath, ME.message);
            end
        end
    end

    if isempty(allStationDaily)
        error("Hic gunluk istasyon verisi uretilmedi.");
    end

    % Şehir bazlı birleştirme
    allStationDaily = allStationDaily(~isnan(allStationDaily.PM10_d_ugm3), :);

    [G2, regU, cityU, dateU, yearU] = findgroups( ...
        allStationDaily.Region, ...
        allStationDaily.City, ...
        allStationDaily.Date, ...
        allStationDaily.Year);

    cityMeanPM10 = splitapply(@mean, allStationDaily.PM10_d_ugm3, G2);
    cityNumStations = splitapply(@(x) numel(unique(x)), allStationDaily.Station, G2);

    CityDailyC = table();
    CityDailyC.Region = regU;
    CityDailyC.City = cityU;
    CityDailyC.Date = dateU;
    CityDailyC.Year = yearU;
    CityDailyC.PM10_d_ugm3 = cityMeanPM10;
    CityDailyC.C_d_gm3 = CityDailyC.PM10_d_ugm3 * 1e-6;
    CityDailyC.N_stations = cityNumStations;

    CityDailyC = sortrows(CityDailyC, {'Region','City','Date'});

    % m_t hesabı
    CityDailyMt = table();
    CityDailyMt.Region = CityDailyC.Region;
    CityDailyMt.City = CityDailyC.City;
    CityDailyMt.Date = CityDailyC.Date;
    CityDailyMt.Year = CityDailyC.Year;
    CityDailyMt.C_t_gm3 = CityDailyC.C_d_gm3;
    CityDailyMt.v_eff_ms = repmat(v_eff_ms, height(CityDailyC), 1);
    CityDailyMt.t_s = repmat(t_s, height(CityDailyC), 1);
    CityDailyMt.theta_deg = repmat(theta_deg, height(CityDailyC), 1);
    CityDailyMt.cos_theta = repmat(cos_theta, height(CityDailyC), 1);
    CityDailyMt.m_t_gm2 = ...
        CityDailyMt.v_eff_ms .* ...
        CityDailyMt.C_t_gm3 .* ...
        CityDailyMt.t_s .* ...
        CityDailyMt.cos_theta;

    CityDailyMt = sortrows(CityDailyMt, {'Region','City','Date'});

    % Excel yaz
    out1 = fullfile(outputsDir, sprintf("CityDailyC_%d.xlsx", targetYear));
    out2 = fullfile(outputsDir, sprintf("CityDailyMt_%d.xlsx", targetYear));

    writetable(CityDailyC, out1, "FileType", "spreadsheet");
    writetable(CityDailyMt, out2, "FileType", "spreadsheet");

    fprintf("Kaydedildi: %s\n", out1);
    fprintf("Kaydedildi: %s\n", out2);
end

function [city, station] = parse_city_station(filename)
    [~, name, ~] = fileparts(filename);
    parts = split(string(name), "_");

    if numel(parts) >= 2
        city = lower(strtrim(parts(1)));
        station = lower(strtrim(strjoin(parts(2:end), "_")));
    else
        city = lower(strtrim(string(name)));
        station = "1";
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