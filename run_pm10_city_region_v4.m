function [StationSeasonTable, CitySeasonTable, RegionSeasonTable] = run_pm10_city_region_v4(baseFolder, targetYear)
% Robust for your Excel layout:
% Row1: "Tarih" in col A
% Row2: "PM10 ( µg/m3 )" in col B
% Data starts row3 with datetime + numeric(string with comma)
%
% Output:
% StationSeasonTable: Region City Station Season Year PM10_s_ugm3 C_s_gm3 N_days N_hours
% CitySeasonTable   : Region City Season Year PM10_s_ugm3 C_s_gm3 N_stations ...
% RegionSeasonTable : Region Season Year PM10_s_ugm3 C_s_gm3 N_cities

    fprintf("Running: %s\n", mfilename('fullpath'));

    if nargin < 1 || strlength(string(baseFolder))==0
        baseFolder = pwd;
    end
    if nargin < 2 || isempty(targetYear)
        targetYear = 2025;
    end

    baseFolder = char(baseFolder);
    if ~isfolder(baseFolder)
        error("baseFolder not found: %s", baseFolder);
    end

    outDir = fullfile(baseFolder, "outputs");
    if ~isfolder(outDir), mkdir(outDir); end

    regionDirs = dir(baseFolder);
    regionDirs = regionDirs([regionDirs.isdir]);
    regionDirs = regionDirs(~ismember({regionDirs.name},{'.','..','outputs'}));

    StationSeasonTable = table();

    for r = 1:numel(regionDirs)
        regionName = string(regionDirs(r).name);
        regionPath = fullfile(baseFolder, regionDirs(r).name);

        files = dir(fullfile(regionPath, "*.xlsx"));
        files = files(~startsWith({files.name}, "~$"));

        for f = 1:numel(files)
            fileName = string(files(f).name);
            filePath = fullfile(regionPath, files(f).name);

            try
                [city, station] = parse_city_station(fileName);

                [dt, pm10] = read_pm10_hourly_robust(filePath);

                rows = station_season_from_hourly(dt, pm10, targetYear, regionName, city, station);
                if ~isempty(rows)
                    StationSeasonTable = [StationSeasonTable; rows]; %#ok<AGROW>
                end

            catch ME
                fprintf("[ERROR] %s -> %s\n", filePath, ME.message);
            end
        end
    end

    if isempty(StationSeasonTable)
        error("No valid station-season rows produced. Check year=%d exists in your data.", targetYear);
    end

    StationSeasonTable = sortrows(StationSeasonTable, {'Region','City','Station','Year','Season'});

    CitySeasonTable = aggregate_city(StationSeasonTable);
    RegionSeasonTable = aggregate_region(CitySeasonTable);

    writetable(StationSeasonTable, fullfile(outDir, "StationSeasonTable.xlsx"));
    writetable(CitySeasonTable,   fullfile(outDir, "CitySeasonTable.xlsx"));
    writetable(RegionSeasonTable, fullfile(outDir, "RegionSeasonTable.xlsx"));

    fprintf("✅ Done. Outputs: %s\n", outDir);
end

% ============================================================
% READ EXCEL (handles split headers like your sample)
% ============================================================
function [dt, pm10] = read_pm10_hourly_robust(filePath)
    sheets = get_sheet_names(filePath);

    for si = 1:numel(sheets)
        sh = sheets(si);

        preview = readcell(filePath, "Sheet", sh, "Range", "A1:AZ3000");
        preview = forceCellMatrix(preview);

        % Find "Tarih" anywhere AND "PM10" anywhere (can be different rows)
        [dtRow, dtCol] = find_token_anywhere(preview, ["tarih","datetime","date","time","zaman"]);
        [pmRow, pmCol] = find_token_anywhere(preview, ["pm10"]);

        if dtRow > 0 && pmRow > 0
            startRow = max(dtRow, pmRow) + 1;
            % read long columns
            [dt, pm10] = read_two_columns(filePath, sh, dtCol, pmCol, startRow);
            if ~isempty(dt), return; end
        end

        % Fallback: content detection (datetime-like + numeric-like)
        [startRow2, dtCol2, pmCol2] = detect_cols_by_content(preview);
        if startRow2 > 0
            [dt, pm10] = read_two_columns(filePath, sh, dtCol2, pmCol2, startRow2);
            if ~isempty(dt), return; end
        end
    end

    error("Could not detect (Datetime + PM10) in any sheet: %s", filePath);
end

function [dt, pm10] = read_two_columns(filePath, sheet, dtCol, pmCol, startRow)
    maxRows = 200000;

    dtL = colnum2excel(dtCol);
    pmL = colnum2excel(pmCol);

    rangeDT = sprintf("%s%d:%s%d", dtL, startRow, dtL, startRow+maxRows);
    rangePM = sprintf("%s%d:%s%d", pmL, startRow, pmL, startRow+maxRows);

    dtRaw = readcell(filePath, "Sheet", sheet, "Range", rangeDT);
    pmRaw = readcell(filePath, "Sheet", sheet, "Range", rangePM);

    dtRaw = forceCellVector(dtRaw);
    pmRaw = forceCellVector(pmRaw);

    n = min(numel(dtRaw), numel(pmRaw));
    dtRaw = dtRaw(1:n);
    pmRaw = pmRaw(1:n);

    dt = to_datetime(dtRaw);
    pm10 = to_numeric(pmRaw);

    k = ~isnat(dt) & ~isnan(pm10);
    dt = dt(k);
    pm10 = pm10(k);
end

function [row, col] = find_token_anywhere(C, tokens)
    row = 0; col = 0;
    [nr, nc] = size(C);
    for r = 1:nr
        for c = 1:nc
            t = norm_token(C{r,c});
            for k = 1:numel(tokens)
                if contains(t, tokens(k))
                    row = r; col = c;
                    return;
                end
            end
        end
    end
end

function [startRow, dtCol, pmCol] = detect_cols_by_content(C)
    startRow = 0; dtCol = 0; pmCol = 0;
    [nr, nc] = size(C);

    % datetime-like column score
    bestDtScore = 0; bestDtCol = 0;
    for c = 1:nc
        colCells = C(:,c);
        dtTry = to_datetime(colCells);
        nonEmpty = ~cellfun(@is_empty_cell_scalar, colCells);
        denom = max(1, sum(nonEmpty));
        score = sum((~isnat(dtTry)) & nonEmpty) / denom;
        if score > bestDtScore
            bestDtScore = score;
            bestDtCol = c;
        end
    end
    if bestDtScore < 0.15
        return;
    end

    bestPmScore = 0; bestPmCol = 0;
    for c = 1:nc
        if c == bestDtCol, continue; end
        colCells = C(:,c);
        xTry = to_numeric(colCells);
        nonEmpty = ~cellfun(@is_empty_cell_scalar, colCells);
        denom = max(1, sum(nonEmpty));
        score = sum((~isnan(xTry)) & nonEmpty) / denom;
        if score > bestPmScore
            bestPmScore = score;
            bestPmCol = c;
        end
    end
    if bestPmScore < 0.15
        return;
    end

    dtTry = to_datetime(C(:,bestDtCol));
    xTry  = to_numeric(C(:,bestPmCol));
    good = ~isnat(dtTry) & ~isnan(xTry);

    idx = find(good, 1, "first");
    if isempty(idx), return; end

    startRow = idx;
    dtCol = bestDtCol;
    pmCol = bestPmCol;
end

% ============================================================
% Station hourly -> daily -> seasonal
% ============================================================
function outT = station_season_from_hourly(dt, pm10_ugm3, targetYear, region, city, station)
    valid = ~isnat(dt) & ~isnan(pm10_ugm3);
    dt = dt(valid);
    pm10_ugm3 = pm10_ugm3(valid);

    if isempty(dt)
        outT = table(); return;
    end

    [dt, pm10_ugm3] = deduplicate_datetime_mean(dt, pm10_ugm3);

    day = dateshift(dt, 'start', 'day');
    [Gd, dayU] = findgroups(day);

    PM10d = splitapply(@(x) mean(x,'omitnan'), pm10_ugm3, Gd);
    Nh_d  = splitapply(@(x) sum(~isnan(x)),   pm10_ugm3, Gd);

    keepDay = (Nh_d > 0) & ~isnan(PM10d);
    dayU  = dayU(keepDay);
    PM10d = PM10d(keepDay);
    Nh_d  = Nh_d(keepDay);

    if isempty(dayU)
        outT = table(); return;
    end

    [seasonDay, seasonYearDay] = datetime_to_season(dayU);

    keepY = (seasonYearDay == targetYear);
    seasonDay = seasonDay(keepY);
    PM10d = PM10d(keepY);
    Nh_d  = Nh_d(keepY);

    if isempty(PM10d)
        outT = table(); return;
    end

    outT = table();
    seasons = categorical(["DJF","MAM","JJA","SON"], ["DJF","MAM","JJA","SON"], 'Ordinal', true);

    for s = ["DJF","MAM","JJA","SON"]
        sCat = categorical(s, categories(seasons), 'Ordinal', true);
        idx = (seasonDay == sCat);
        if ~any(idx), continue; end

        pm_s = mean(PM10d(idx), 'omitnan');   % ug/m3
        Nd_s = sum(~isnan(PM10d(idx)));
        Nh_s = sum(Nh_d(idx));

        C_s = pm_s * 1e-6;                   % g/m3

        outT = [outT; table( ...
            string(region), string(city), string(station), sCat, double(targetYear), ...
            pm_s, C_s, Nd_s, Nh_s, ...
            'VariableNames', {'Region','City','Station','Season','Year','PM10_s_ugm3','C_s_gm3','N_days','N_hours'})]; %#ok<AGROW>
    end
end

function CitySeasonTable = aggregate_city(StationSeasonTable)
    [G, reg, city, seas, yr] = findgroups( ...
        StationSeasonTable.Region, StationSeasonTable.City, StationSeasonTable.Season, StationSeasonTable.Year);

    PM10_city = splitapply(@(x) mean(x,'omitnan'), StationSeasonTable.PM10_s_ugm3, G);
    C_city    = PM10_city * 1e-6;

    Nst = splitapply(@(x) numel(unique(string(x))), StationSeasonTable.Station, G);

    CitySeasonTable = table(reg, city, seas, yr, PM10_city, C_city, Nst, ...
        'VariableNames', {'Region','City','Season','Year','PM10_s_ugm3','C_s_gm3','N_stations'});
end

function RegionSeasonTable = aggregate_region(CitySeasonTable)
    [G, reg, seas, yr] = findgroups(CitySeasonTable.Region, CitySeasonTable.Season, CitySeasonTable.Year);

    PM10_reg = splitapply(@(x) mean(x,'omitnan'), CitySeasonTable.PM10_s_ugm3, G);
    C_reg    = PM10_reg * 1e-6;

    Ncity = splitapply(@(x) numel(unique(string(x))), CitySeasonTable.City, G);

    RegionSeasonTable = table(reg, seas, yr, PM10_reg, C_reg, Ncity, ...
        'VariableNames', {'Region','Season','Year','PM10_s_ugm3','C_s_gm3','N_cities'});
end

% ============================================================
% Seasons: DJF year rule
% ============================================================
function [seasonCode, seasonYear] = datetime_to_season(dt)
    m = month(dt);
    y = year(dt);

    seasonYear = y;
    sc = strings(size(dt));

    idxDJF = (m==12) | (m==1) | (m==2);
    sc(idxDJF) = "DJF";
    seasonYear(m==12) = y(m==12) + 1;

    sc((m==3)|(m==4)|(m==5)) = "MAM";
    sc((m==6)|(m==7)|(m==8)) = "JJA";
    sc((m==9)|(m==10)|(m==11)) = "SON";

    seasonCode = categorical(sc, ["DJF","MAM","JJA","SON"], 'Ordinal', true);
end

% ============================================================
% Filename parse: city first token, station rest
% ============================================================
function [city, station] = parse_city_station(fileName)
    name = erase(string(fileName), ".xlsx");
    name = replace(name, "_", " ");
    name = regexprep(name, "\s+", " ");
    name = strtrim(name);

    parts = split(lower(name), " ");
    parts(parts=="") = [];

    if numel(parts) == 0
        city = "unknown_city";
        station = "unknown_station";
    elseif numel(parts) == 1
        city = parts(1);
        station = "unknown_station";
    else
        city = parts(1);
        station = strjoin(parts(2:end), " ");
    end
end

% ============================================================
% Helpers: FORCE CELL (kills brace indexing issue permanently)
% ============================================================
function C = forceCellMatrix(X)
    if iscell(X)
        C = X; return;
    end
    C = cell(size(X));
    for i = 1:numel(X)
        C{i} = X(i);
    end
end

function v = forceCellVector(X)
    if iscell(X)
        v = X(:); return;
    end
    v = cell(numel(X),1);
    for i = 1:numel(X)
        v{i} = X(i);
    end
end

% ============================================================
% Converters (your PM10 uses comma)
% ============================================================
function dt = to_datetime(xCell)
    xCell = forceCellVector(xCell);
    dt = NaT(size(xCell));

    for i = 1:numel(xCell)
        xi = xCell{i};
        try
            if isdatetime(xi)
                dt(i) = xi;
            elseif isnumeric(xi) && isscalar(xi) && ~isnan(xi)
                dt(i) = datetime(xi, 'ConvertFrom', 'excel');
            else
                s = string(xi);
                s = strip(s);
                if strlength(s)==0
                    dt(i) = NaT;
                else
                    % try TR formats first
                    try
                        dt(i) = datetime(s, 'InputFormat', 'dd.MM.yyyy HH:mm:ss', 'Locale','tr_TR');
                    catch
                        try
                            dt(i) = datetime(s, 'InputFormat', 'dd.MM.yyyy HH:mm', 'Locale','tr_TR');
                        catch
                            dt(i) = datetime(s, 'Locale','tr_TR');
                        end
                    end
                end
            end
        catch
            dt(i) = NaT;
        end
    end
end

function x = to_numeric(xCell)
    xCell = forceCellVector(xCell);
    x = nan(size(xCell));

    for i = 1:numel(xCell)
        xi = xCell{i};
        try
            if isnumeric(xi) && isscalar(xi)
                x(i) = double(xi);
            else
                s = string(xi);
                s = strip(s);
                s = replace(s, ",", ".");
                if strlength(s)==0
                    x(i) = NaN;
                else
                    x(i) = str2double(s);
                end
            end
        catch
            x(i) = NaN;
        end
    end
end

function tf = is_empty_cell_scalar(v)
    if isempty(v), tf = true; return; end
    try
        m = ismissing(v);
        if isscalar(m)
            if m, tf = true; return; end
        else
            if all(m(:)), tf = true; return; end
        end
    catch
    end
    if ischar(v)
        tf = isempty(strtrim(v)); return;
    end
    if isstring(v)
        tf = (strlength(strtrim(v))==0); return;
    end
    tf = false;
end

function s = norm_token(x)
    s = string(x);
    s = lower(strtrim(s));
    s = replace(s, newline, " ");
    s = replace(s, char(160), " ");
    s = replace(s, "µ", "u");
    s = replace(s, "³", "3");
    s = regexprep(s, "[^a-z0-9ığüşöç]", "");
end

function sheets = get_sheet_names(filePath)
    try
        sheets = sheetnames(filePath);
    catch
        [~, sh] = xlsfinfo(filePath);
        sheets = string(sh);
    end
end

function [t2, x2] = deduplicate_datetime_mean(t, x)
    [t, ord] = sort(t);
    x = x(ord);
    [G, tu] = findgroups(t);
    xu = splitapply(@(v) mean(v,'omitnan'), x, G);
    t2 = tu;
    x2 = xu;
end

function L = colnum2excel(n)
    L = "";
    while n > 0
        r = mod(n-1, 26);
        L = char(r + double('A')) + L;
        n = floor((n-1)/26);
    end
    L = char(L);
end
