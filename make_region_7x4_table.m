function [PM10wide, Cwide, outFile] = make_region_7x4_table(outputsDir, targetYear)
%MAKE_REGION_7X4_TABLE  Robust 7×4 Region×Season table from outputs folder
%
% INPUTS
%   outputsDir : "... \hava kirliliği verileri\outputs"
%   targetYear : e.g., 2025
%
% READS
%   outputsDir\RegionSeasonTable.xlsx
%
% WRITES
%   outputsDir\Region_7x4_<year>.xlsx

    if nargin < 1 || strlength(string(outputsDir)) == 0
        error("outputsDir vermelisin. Örn: C:\...\hava kirliliği verileri\outputs");
    end
    if nargin < 2 || isempty(targetYear)
        targetYear = 2025;
    end

    outputsDir = char(outputsDir);
    inFile = fullfile(outputsDir, "RegionSeasonTable.xlsx");
    if ~isfile(inFile)
        error("Bulunamadı: %s (önce run_pm10_city_region_v4 çalışmalı)", inFile);
    end

    seasonOrder = ["DJF","MAM","JJA","SON"];

    T = readtable(inFile);

    % ---- required columns ----
    need = ["Region","Season","Year","PM10_s_ugm3","C_s_gm3"];
    miss = setdiff(need, string(T.Properties.VariableNames));
    if ~isempty(miss)
        error("RegionSeasonTable.xlsx eksik kolon(lar): %s", strjoin(miss, ", "));
    end

    % ---- normalize ----
    T.Region = strtrim(string(T.Region));
    T.Season = strtrim(string(T.Season));
    T.Year   = double(T.Year);

    T.PM10_s_ugm3 = forceNumeric(T.PM10_s_ugm3);
    T.C_s_gm3     = forceNumeric(T.C_s_gm3);

    T.Season = normalizeSeasonToDJF(T.Season);

    % ---- drop invalid grouping rows ----
    bad = (strlength(T.Region)==0) | (strlength(T.Season)==0) | isnan(T.Year);
    T(bad,:) = [];

    % ---- filter year ----
    T = T(T.Year == targetYear, :);
    if isempty(T)
        error("targetYear=%d için geçerli satır yok.", targetYear);
    end

    % ---- average duplicates: findgroups+splitapply ----
    [G, reg, seas, yr] = findgroups(T.Region, T.Season, T.Year);
    pmMean = splitapply(@(x) mean(x,'omitnan'), T.PM10_s_ugm3, G);
    cMean  = splitapply(@(x) mean(x,'omitnan'), T.C_s_gm3,     G);

    T2 = table(reg, seas, yr, pmMean, cMean, ...
        'VariableNames', {'Region','Season','Year','PM10_s_ugm3','C_s_gm3'});

    % ---- ordered season ----
    T2.Season = categorical(string(T2.Season), seasonOrder, "Ordinal", true);

    % ===== PM10 pivot (NO curly/brace selection) =====
    TPM = table(string(T2.Region), T2.Season, T2.PM10_s_ugm3, ...
        'VariableNames', {'Region','Season','PM10_s_ugm3'});
    PM10wide = unstack(TPM, 'PM10_s_ugm3', 'Season');
    PM10wide = ensureSeasonCols(PM10wide, seasonOrder);
    PM10wide = PM10wide(:, ["Region", seasonOrder]);
    PM10wide{:,seasonOrder} = round(PM10wide{:,seasonOrder}, 2);

    % ===== C pivot =====
    TC = table(string(T2.Region), T2.Season, T2.C_s_gm3, ...
        'VariableNames', {'Region','Season','C_s_gm3'});
    Cwide = unstack(TC, 'C_s_gm3', 'Season');
    Cwide = ensureSeasonCols(Cwide, seasonOrder);
    Cwide = Cwide(:, ["Region", seasonOrder]);
    Cwide{:,seasonOrder} = round(Cwide{:,seasonOrder}, 9);

    disp("=== Region × Season : PM10_s (µg/m³) ==="); disp(PM10wide);
    disp("=== Region × Season : C_s (g/m³) ===");     disp(Cwide);

    outFile = fullfile(outputsDir, sprintf("Region_7x4_%d.xlsx", targetYear));
    writetable(PM10wide, outFile, "Sheet", "PM10_ugm3");
    writetable(Cwide,    outFile, "Sheet", "C_gm3");
    fprintf("Kaydedildi: %s\n", outFile);
end

% ================= helpers =================

function W = ensureSeasonCols(W, seasonOrder)
    % make sure columns exist even if a season is missing
    for s = seasonOrder
        if ~ismember(s, string(W.Properties.VariableNames))
            W.(s) = nan(height(W),1);
        end
    end
end

function sOut = normalizeSeasonToDJF(sIn)
    s = lower(string(sIn));
    s = strtrim(s);

    s2 = replace(s, "ı", "i");
    s2 = replace(s2, "ş", "s");
    s2 = replace(s2, "ğ", "g");
    s2 = replace(s2, "ç", "c");
    s2 = replace(s2, "ö", "o");
    s2 = replace(s2, "ü", "u");

    sOut = strings(size(s2));

    sOut(s2=="djf") = "DJF";
    sOut(s2=="mam") = "MAM";
    sOut(s2=="jja") = "JJA";
    sOut(s2=="son") = "SON";

    sOut(ismember(s2, ["kis","winter"])) = "DJF";
    sOut(ismember(s2, ["ilkbahar","spring"])) = "MAM";
    sOut(ismember(s2, ["yaz","summer"])) = "JJA";
    sOut(ismember(s2, ["sonbahar","autumn","fall"])) = "SON";

    idx = (strlength(sOut)==0);
    sOut(idx) = upper(string(sIn(idx)));
end

function x = forceNumeric(x)
    if isnumeric(x)
        x = double(x); return;
    end
    xs = string(x);
    xs = strtrim(xs);
    xs = replace(xs, ",", ".");
    x = str2double(xs);
end