namespace NetworkAdapterSwitcher;

internal enum Word
{
    Subtitle, Wifi, Adapters, Refresh, Log, Settings, Theme, Dark, Light, Language,
    Interface, Active, FirstNetwork, SecondNetwork, FirstAdapter, SecondAdapter,
    Virtual, Bluetooth, Admin, Elevate, PickWifi, PickAdapters, Disconnect, Unknown,
    Saved, Enabled, Disabled, Connected, Connect, Switch, Busy, WifiHint, AdapterHint, SystemTheme, EmptyLog, CopyLog, LogCopied
}

internal static class UiWords
{
    public static readonly (string Key, string Label)[] Languages =
    [
        ("English", "English"), ("Russian", "Русский"), ("French", "Français"),
        ("German", "Deutsch"), ("Chinese", "中文"), ("Polish", "Polski"), ("Japanese", "日本語")
    ];

    private static readonly string[][] Texts =
    [
        ("Switch networks and adapters|Wi-Fi networks|Adapters|Refresh|Activity|Settings|Theme|Dark|Light|Language|" +
         "Wi-Fi adapter|Active connection|Network A|Network B|Adapter A|Adapter B|Virtual|Bluetooth|Administrator mode|Run as administrator|" +
         "Choose two different saved networks.|Choose two different adapters.|Disconnected|Profile unavailable|" +
         "Saved profile|Enabled|Disabled|Connected|Connect to|Switch to|Working…|" +
         "Choose networks already saved in Windows.|The target link is checked before the active adapter is disabled.|Follow Windows|No activity yet.|Copy log|Copied").Split('|'),
        ("Переключение сетей и адаптеров|Сети Wi-Fi|Адаптеры|Обновить|Журнал|Настройки|Тема|Тёмная|Светлая|Язык|" +
         "Адаптер Wi-Fi|Активное подключение|Сеть A|Сеть B|Адаптер A|Адаптер B|Виртуальные|Bluetooth|Режим администратора|Запустить от администратора|" +
         "Выберите две разные сохранённые сети.|Выберите два разных адаптера.|Нет подключения|Имя профиля недоступно|" +
         "Сохранённый профиль|Включён|Отключён|Подключено|Подключить|Переключить на|Выполняется…|" +
         "Выберите сети, уже сохранённые в Windows.|Перед отключением активного адаптера проверяется новое подключение.|Как в Windows|Событий пока нет.|Копировать|Скопировано").Split('|'),
        ("Changer de réseau et d’adaptateur|Réseaux Wi-Fi|Adaptateurs|Actualiser|Journal|Paramètres|Thème|Sombre|Clair|Langue|" +
         "Adaptateur Wi-Fi|Connexion active|Réseau A|Réseau B|Adaptateur A|Adaptateur B|Virtuels|Bluetooth|Mode administrateur|Exécuter en administrateur|" +
         "Choisissez deux réseaux enregistrés différents.|Choisissez deux adaptateurs différents.|Déconnecté|Profil indisponible|" +
         "Profil enregistré|Activé|Désactivé|Connecté|Connecter|Basculer vers|En cours…|" +
         "Choisissez des réseaux enregistrés dans Windows.|La nouvelle connexion est vérifiée avant la désactivation de l’adaptateur actif.|Selon Windows|Aucune activité.|Copier|Copié").Split('|'),
        ("Netzwerke und Adapter wechseln|WLAN-Netze|Adapter|Aktualisieren|Protokoll|Einstellungen|Design|Dunkel|Hell|Sprache|" +
         "WLAN-Adapter|Aktive Verbindung|Netz A|Netz B|Adapter A|Adapter B|Virtuell|Bluetooth|Administratormodus|Als Administrator starten|" +
         "Zwei verschiedene gespeicherte Netze auswählen.|Zwei verschiedene Adapter auswählen.|Getrennt|Profil nicht verfügbar|" +
         "Gespeichertes Profil|Aktiviert|Deaktiviert|Verbunden|Verbinden mit|Wechseln zu|Wird ausgeführt…|" +
         "In Windows gespeicherte Netze auswählen.|Die neue Verbindung wird geprüft, bevor der aktive Adapter deaktiviert wird.|Wie in Windows|Keine Einträge.|Kopieren|Kopiert").Split('|'),
        ("切换网络和适配器|Wi-Fi 网络|适配器|刷新|日志|设置|主题|深色|浅色|语言|" +
         "Wi-Fi 适配器|当前连接|网络 A|网络 B|适配器 A|适配器 B|虚拟|蓝牙|管理员模式|以管理员身份运行|" +
         "请选择两个不同的已保存网络。|请选择两个不同的适配器。|未连接|无法读取配置文件|" +
         "已保存配置文件|已启用|已禁用|已连接|连接到|切换到|处理中…|" +
         "请选择 Windows 中已保存的网络。|禁用当前适配器前将检查新连接。|跟随 Windows|暂无活动。|复制日志|已复制").Split('|'),
        ("Przełączanie sieci i kart|Sieci Wi-Fi|Karty sieciowe|Odśwież|Dziennik|Ustawienia|Motyw|Ciemny|Jasny|Język|" +
         "Karta Wi-Fi|Aktywne połączenie|Sieć A|Sieć B|Karta A|Karta B|Wirtualne|Bluetooth|Tryb administratora|Uruchom jako administrator|" +
         "Wybierz dwie różne zapisane sieci.|Wybierz dwie różne karty.|Rozłączono|Profil niedostępny|" +
         "Zapisany profil|Włączona|Wyłączona|Połączono|Połącz z|Przełącz na|Praca w toku…|" +
         "Wybierz sieci zapisane w systemie Windows.|Nowe połączenie jest sprawdzane przed wyłączeniem aktywnej karty.|Jak w Windows|Brak zdarzeń.|Kopiuj|Skopiowano").Split('|'),
        ("ネットワークとアダプターを切り替え|Wi-Fi ネットワーク|アダプター|更新|履歴|設定|テーマ|ダーク|ライト|言語|" +
         "Wi-Fi アダプター|現在の接続|ネットワーク A|ネットワーク B|アダプター A|アダプター B|仮想|Bluetooth|管理者モード|管理者として実行|" +
         "保存済みの異なるネットワークを2つ選択します。|異なるアダプターを2つ選択します。|未接続|プロファイルを取得できません|" +
         "保存済みプロファイル|有効|無効|接続済み|接続先|切り替え先|処理中…|" +
         "Windows に保存済みのネットワークを選択します。|現在のアダプターを無効にする前に、新しい接続を確認します。|Windows に合わせる|履歴はありません。|コピー|コピー済み").Split('|')
    ];

    public static string Get(string language, Word word)
    {
        var index = Array.FindIndex(Languages, item => item.Key == language);
        return Texts[Math.Max(0, index)][(int)word];
    }
}
