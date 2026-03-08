function CityDailyOmega = compute_city_daily_omega_v1(outputsFolder, targetYear)
% compute_city_daily_omega_v1
% Gunluk m_t verisinden kumulatif birikim omega_t hesaplar
%
% Formul:
%   omega_t = omega_(t-1) + m_t
%
% Girdi:
%   outputsFolder = "C:\Users\User\Desktop\2. kısım\outputs";
%   targetYear    = 2025
%
% Beklenen dosya:
%   CityDailyMt_2025.xlsx
%
% Cikti:
%   CityDailyOmega_2025.xlsx

    arguments
        outputsFolder (1,1) string
        targetYear (1,1) double
    end

    % Girdi dosyasi
    inFile = fullfile(outputsFolder, sprintf("CityDailyMt_%d.xlsx", targetYear));

    if ~exist(inFile, "file")
        error("Girdi dosyasi bulunamadi: %s", inFile);
    end

    % Excel oku
    T = readtable(inFile, "PreserveVariableNames", true);

    % Gerekli kolonlar
    requiredVars = {'Region','City','Date','Year','m_t_gm2'};
    for i = 1:numel(requiredVars)
        if ~ismember(requiredVars{i}, T.Properties.VariableNames)
            error("Gerekli kolon eksik: %s", requiredVars{i});
        end
    end

    % Date kolonu datetime degilse donustur
    if ~isdatetime(T.Date)
        try
            T.Date = datetime(T.Date);
        catch
            error("Date kolonu datetime formatina cevrilemedi.");
        end
    end

    % Siralama
    T = sortrows(T, {'Region','City','Date'});

    % Sehir bazinda kumulatif toplam
    CityDailyOmega = T;
    CityDailyOmega.omega_t_gm2 = zeros(height(T), 1);

    [G, ~, ~] = findgroups(T.Region, T.City);

    for g = 1:max(G)
        idx = find(G == g);

        % Her sehir icin tarih sirasi garanti
        [~, order] = sort(T.Date(idx));
        idx = idx(order);

        mtVals = T.m_t_gm2(idx);
        omegaVals = cumsum(mtVals);

        CityDailyOmega.omega_t_gm2(idx) = omegaVals;
    end

    % Kolon duzeni
    keepVars = {'Region','City','Date','Year','C_t_gm3','v_eff_ms','t_s','theta_deg','cos_theta','m_t_gm2','omega_t_gm2'};
    keepVars = keepVars(ismember(keepVars, CityDailyOmega.Properties.VariableNames));
    CityDailyOmega = CityDailyOmega(:, keepVars);

    % Excel yaz
    outFile = fullfile(outputsFolder, sprintf("CityDailyOmega_%d.xlsx", targetYear));
    writetable(CityDailyOmega, outFile, "FileType", "spreadsheet");

    fprintf("Kaydedildi: %s\n", outFile);
end