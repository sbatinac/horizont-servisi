/* Katalog za Silicon Graveyard / Tech Outlet: kategorije, proizvođači i tipovi (koristi ga admin i web stranica) */
window.HZ_KATALOG = {
  'Procesor (CPU)': { pro: ['Intel', 'AMD'], tip: ['LGA775', 'LGA1150', 'LGA1151', 'LGA1155', 'LGA1200', 'LGA1700', 'LGA2011', 'AM3/AM3+', 'AM4', 'AM5', 'FM2/FM2+', 'Laptop (BGA)'], spec: 'npr. Core i5-8400, 6 jezgri, 2,8 GHz' },
  'Matična ploča': { pro: ['ASUS', 'MSI', 'Gigabyte', 'ASRock', 'Biostar', 'Intel', 'Dell', 'HP', 'Lenovo'], tip: ['ATX', 'Micro-ATX', 'Mini-ITX', 'E-ATX', 'Laptop'], spec: 'npr. B450, socket AM4, DDR4, M.2' },
  'Radna memorija (RAM)': { pro: ['Kingston', 'Corsair', 'Crucial', 'G.Skill', 'Samsung', 'SK hynix', 'Micron', 'ADATA', 'Patriot', 'TeamGroup', 'HyperX'], tip: ['DDR2', 'DDR3', 'DDR3L', 'DDR4', 'DDR5', 'SO-DIMM DDR3', 'SO-DIMM DDR4', 'SO-DIMM DDR5'], spec: 'npr. 2 × 8 GB, 3200 MHz, CL16' },
  'Grafička kartica (GPU)': { pro: ['NVIDIA', 'AMD', 'Intel', 'ASUS', 'MSI', 'Gigabyte', 'Sapphire', 'PowerColor', 'Palit', 'ZOTAC', 'EVGA', 'Gainward', 'XFX'], tip: ['PCIe x16', 'Niskoprofilna', 'Laptop (MXM)'], spec: 'npr. GTX 1060 6 GB, 1 × HDMI, 3 × DP' },
  'Tvrdi disk (HDD)': { pro: ['Western Digital', 'Seagate', 'Toshiba', 'Hitachi / HGST', 'Samsung'], tip: ['3,5" SATA', '2,5" SATA', 'Vanjski USB'], spec: 'npr. 1 TB, 7200 o/min, zdravlje 98 %' },
  'SSD / NVMe': { pro: ['Samsung', 'Kingston', 'Crucial', 'Western Digital', 'SanDisk', 'Intel', 'ADATA', 'Patriot', 'Lexar', 'Transcend'], tip: ['2,5" SATA', 'M.2 SATA', 'M.2 NVMe', 'mSATA', 'Vanjski USB'], spec: 'npr. 512 GB, zdravlje 95 %' },
  'Napajanje (PSU)': { pro: ['Corsair', 'Seasonic', 'be quiet!', 'Cooler Master', 'EVGA', 'Chieftec', 'Thermaltake', 'FSP', 'LC Power', 'Laptop punjač'], tip: ['ATX', 'SFX', 'TFX', 'Laptop punjač', 'Modularno', 'Polumodularno'], spec: 'npr. 650 W, 80+ Bronze' },
  'Kućište': { pro: ['Fractal Design', 'NZXT', 'Corsair', 'be quiet!', 'Cooler Master', 'Lian Li', 'Phanteks', 'Zalman', 'Sharkoon', 'Genesis'], tip: ['Full Tower', 'Midi Tower', 'Mini Tower', 'Mini-ITX', 'Desktop / SFF'], spec: 'npr. 3 ventilatora, staklena stranica' },
  'Hlađenje': { pro: ['Noctua', 'be quiet!', 'Cooler Master', 'Arctic', 'DeepCool', 'Corsair', 'NZXT', 'Intel (stock)', 'AMD (stock)'], tip: ['CPU hladnjak (zračni)', 'AIO vodeno', 'Ventilator kućišta', 'Termalna pasta', 'Laptop hladnjak'], spec: 'npr. 120 mm, socket AM4 / LGA1200' },
  'Mrežna / zvučna / dodatna kartica': { pro: ['TP-Link', 'Intel', 'ASUS', 'Realtek', 'Creative', 'D-Link', 'Netgear'], tip: ['Wi-Fi kartica', 'LAN kartica', 'Zvučna kartica', 'USB kartica', 'Capture kartica', 'RAID kontroler'], spec: 'npr. Wi-Fi 6, Bluetooth 5.2, PCIe' },
  'Laptop dijelovi': { pro: ['Lenovo', 'HP', 'Dell', 'Acer', 'ASUS', 'Toshiba', 'Apple', 'MSI', 'Samsung'], tip: ['Ekran', 'Tipkovnica', 'Baterija', 'Matična ploča', 'Kućište / poklopac', 'Ventilator', 'Punjač', 'Šarke', 'Ostalo'], spec: 'npr. 15,6" FHD, 30 pin eDP' },
  'Konzole i dijelovi': { pro: ['Sony PlayStation', 'Microsoft Xbox', 'Nintendo', 'Valve'], tip: ['Konzola', 'Kontroler', 'Napajanje', 'Ventilator', 'Disk / pogon', 'Kabel', 'Ostalo'], spec: 'npr. PS4 Slim 500 GB, 1 kontroler' },
  'Monitor': { pro: ['Dell', 'LG', 'Samsung', 'AOC', 'ASUS', 'BenQ', 'Philips', 'HP', 'Lenovo', 'iiyama'], tip: ['do 22"', '23–24"', '27"', '32" i više'], spec: 'npr. 24", 1920×1080, 75 Hz, HDMI' },
  'Periferija': { pro: ['Logitech', 'Razer', 'SteelSeries', 'HyperX', 'Microsoft', 'Genius', 'Trust', 'Corsair'], tip: ['Tipkovnica', 'Miš', 'Slušalice', 'Web kamera', 'Zvučnici', 'Gamepad'], spec: 'npr. mehanička, HR raspored' },
  'Kablovi i adapteri': { pro: ['Razno', 'Ugreen', 'Delock', 'Lindy', 'Gembird'], tip: ['HDMI', 'DisplayPort', 'VGA / DVI', 'USB', 'Napajanje', 'SATA', 'Adapter / reduktor', 'Ostalo'], spec: 'npr. HDMI 2.0, 2 m' },
  'Kompletno računalo / laptop': { pro: ['Dell', 'HP', 'Lenovo', 'Acer', 'ASUS', 'Apple', 'MSI', 'Custom'], tip: ['Desktop', 'Laptop', 'Mini PC', 'All-in-One'], spec: 'npr. i5-8400, 16 GB RAM, 512 GB SSD, Windows 11' },
  'Ostalo': { pro: ['Razno'], tip: ['Ostalo'], spec: 'kratko opišite' }
};
window.HZ_STANJA = ['Rabljeno – ispravno, testirano', 'Rabljeno – ispravno', 'Novo – otvoreno pakiranje', 'Novo – zapakirano', 'Neispravno – za dijelove'];
window.HZ_JAMSTVA = ['bez jamstva', '7 dana', '14 dana', '30 dana', '3 mjeseca', '6 mjeseci'];
