/**
 * FILE: setup.gs
 * Fungsi: Membuat struktur Sheet (Produk, Supplier, Pengaturan, Pesanan, Voucher, Ulasan)
 * beserta data awal (dummy data) Hibualamo Store.
 * Cara Pakai: Pilih fungsi "setupDatabase" di editor Google Apps Script lalu klik "Run".
 */

function setupDatabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1. SETUP SHEET SUPPLIER
  let sheetSupplier = ss.getSheetByName('Supplier');
  if (!sheetSupplier) {
    sheetSupplier = ss.insertSheet('Supplier');
  } else {
    sheetSupplier.clear();
  }

  const headerSupplier = [
    'ID_Supplier', 'Nama_Supplier', 'Jenis_Supplier', 'Alamat', 'Kontak_Person'
  ];
  sheetSupplier.appendRow(headerSupplier);
  styleHeader(sheetSupplier, headerSupplier.length);

  const dataSupplier = [
    ['SUP-001', 'PT Sumber Elektronik Jaya', 'Distributor Resmi', 'Jl. Mangga Dua Raya No. 18, Jakarta Pusat', 'Bpk. Hendra (081299887766)'],
    ['SUP-002', 'CV Konveksi Nusantara', 'Produsen Lokal', 'Jl. Cigondewah Kidul No. 42, Kota Bandung', 'Ibu Rina (081322334455)'],
    ['SUP-003', 'UD Pangan Nusantara Mandiri', 'Grosir / Produsen', 'Jl. Soekarno Hatta No. 89, Bandar Lampung', 'Bpk. Anton (081566778899)'],
    ['SUP-004', 'Global Living Importir', 'Importir', 'Kawasan Pergudangan Pantai Indah Dadap, Tangerang', 'Bpk. Kevin (081122339988)']
  ];
  dataSupplier.forEach(row => sheetSupplier.appendRow(row));

  // 2. SETUP SHEET PENGATURAN (Password Admin, Kategori, Logo, No WA Admin, QRIS, Rekening Bank, & Gmail Pemilik)
  let sheetPengaturan = ss.getSheetByName('Pengaturan');
  if (!sheetPengaturan) {
    sheetPengaturan = ss.insertSheet('Pengaturan');
  } else {
    sheetPengaturan.clear();
  }

  const headerPengaturan = ['Kunci', 'Nilai'];
  sheetPengaturan.appendRow(headerPengaturan);
  styleHeader(sheetPengaturan, headerPengaturan.length);

  let defaultEmail = '';
  try {
    defaultEmail = Session.getActiveUser().getEmail() || '';
  } catch (e) {
    defaultEmail = '';
  }

  const dataPengaturan = [
    ['Password_Admin', '123456'],
    ['Kategori', 'Elektronik,Fashion,Makanan,Aksesoris,Rumah Tangga'],
    ['Logo_Url', ''],
    ['No_WA_Admin', '6281234567890'],
    ['QRIS_Url', 'https://placehold.co/300x300/1e293b/ffffff?text=QRIS+HibualamoStore'],
    ['No_Rekening', 'Bank BCA: 8720192837'],
    ['Nama_Pemilik_Rekening', 'Hibualamo Store Official'],
    ['Gmail_Pemilik', defaultEmail]
  ];
  dataPengaturan.forEach(row => sheetPengaturan.appendRow(row));

  // 3. SETUP SHEET PRODUK
  let sheetProduk = ss.getSheetByName('Produk');
  if (!sheetProduk) {
    sheetProduk = ss.insertSheet('Produk');
  } else {
    sheetProduk.clear();
  }

  const headerProduk = [
    'ID_Produk', 'Nama', 'Kategori', 'Harga_Supplier', 'Harga_Asli', 'Harga_Diskon',
    'Stok', 'Terjual', 'Rating', 'URL_Gambar', 'Varian',
    'Is_FlashSale', 'Badge', 'Deskripsi', 'ID_Supplier', 'Nama_Supplier'
  ];
  sheetProduk.appendRow(headerProduk);
  styleHeader(sheetProduk, headerProduk.length);

  const dataProduk = [
    [
      'PRD-001',
      'TWS Earbuds Bluetooth 5.3 Bass Hifi Waterproof IPX5 Gaming Low Latency',
      'Elektronik',
      65000,
      250000,
      99000,
      45,
      1420,
      4.9,
      'https://images.unsplash.com/photo-1590658268037-6bf12165a8df?auto=format&fit=crop&w=600&q=80',
      'Hitam,Putih,Pink Pastel',
      true,
      'Mall',
      'TWS Bluetooth 5.3 generasi terbaru dengan suara Super Bass jernih, baterai tahan hingga 24 jam bersama charging case, dan fitur gaming mode tanpa delay.',
      'SUP-001',
      'PT Sumber Elektronik Jaya'
    ],
    [
      'PRD-002',
      'Kemeja Flanel Pria Lengan Panjang Premium Casual Distro Kotak-Kotak',
      'Fashion',
      58000,
      185000,
      89000,
      80,
      890,
      4.8,
      'https://images.unsplash.com/photo-1602810318383-e386cc2a3ccf?auto=format&fit=crop&w=600&q=80',
      'Ukuran M,Ukuran L,Ukuran XL',
      true,
      'Star+',
      'Bahan katun flanel import tebal, adem, dan tidak mudah berbulu. Cocok untuk kuliah, kerja santai, maupun hangout.',
      'SUP-002',
      'CV Konveksi Nusantara'
    ],
    [
      'PRD-003',
      'Smartwatch Sport AMOLED Jam Tangan Pintar Heart Rate Monitor Waterproof',
      'Elektronik',
      135000,
      450000,
      195000,
      25,
      2150,
      4.9,
      'https://images.unsplash.com/photo-1523275335684-37898b6baf30?auto=format&fit=crop&w=600&q=80',
      'Black Strap,Silver Strap,Rose Gold',
      true,
      'Mall',
      'Layar sentuh penuh HD AMOLED 1.85 inch dengan pemantau detak jantung 24 jam, penghitung langkah, notifikasi WhatsApp, dan tahan air IP68.',
      'SUP-001',
      'PT Sumber Elektronik Jaya'
    ],
    [
      'PRD-004',
      'Sepatu Sneakers Running Olahraga Pria Wanita Breathable Ringan',
      'Fashion',
      95000,
      320000,
      149000,
      60,
      640,
      4.7,
      'https://images.unsplash.com/photo-1542291026-7eec264c27ff?auto=format&fit=crop&w=600&q=80',
      'Size 39,Size 40,Size 41,Size 42,Size 43',
      false,
      'Star+',
      'Desain ergonomis dengan sol karet anti selip dan bahan rajut berpori yang membuat kaki tetap sejuk saat berlari atau jalan jauh.',
      'SUP-002',
      'CV Konveksi Nusantara'
    ],
    [
      'PRD-005',
      'Keripik Pisang Coklat Lumer Lampung Kemasan Pouch 250gr Crispy',
      'Makanan',
      18000,
      45000,
      28000,
      150,
      3400,
      4.9,
      'https://images.unsplash.com/photo-1621996346565-e3d5d6281298?auto=format&fit=crop&w=600&q=80',
      'Coklat Lumer,Matcha,Keju Susu,Tiramisu',
      true,
      'Star+',
      'Keripik pisang kepok pilihan yang renyah dengan balutan bubuk dan saus coklat premium tebal. Tanpa bahan pengawet, halal dan bikin nagih!',
      'SUP-003',
      'UD Pangan Nusantara Mandiri'
    ],
    [
      'PRD-006',
      'Tas Ransel Laptop Backpack Waterproof Anti Maling USB Port 15.6 Inch',
      'Aksesoris',
      75000,
      210000,
      115000,
      40,
      530,
      4.8,
      'https://images.unsplash.com/photo-1553062407-98eeb64c6a62?auto=format&fit=crop&w=600&q=80',
      'Abu-abu,Hitam,Navy',
      false,
      'Mall',
      'Kapasitas besar dengan kompartemen khusus laptop hingga 15.6 inci, bahan Oxford tahan cipratan air, dan dilengkapi port USB eksternal.',
      'SUP-004',
      'Global Living Importir'
    ],
    [
      'PRD-007',
      'Lampu Meja Belajar LED Rechargeable Touch Dimmer Eye Protection',
      'Rumah Tangga',
      38000,
      120000,
      59000,
      95,
      1120,
      4.8,
      'https://images.unsplash.com/photo-1507473885765-e6ed057f782c?auto=format&fit=crop&w=600&q=80',
      'Putih Warm,Putih Cool',
      false,
      'Star+',
      'Lampu baca LED dengan 3 tingkat kecerahan sentuh, leher fleksibel 360 derajat, serta baterai tanam yang dapat diisi ulang via kabel Type-C.',
      'SUP-004',
      'Global Living Importir'
    ],
    [
      'PRD-008',
      'Kopi Arabika Gayo Aceh Specialty 250gr Biji / Bubuk Fresh Roast',
      'Makanan',
      45000,
      95000,
      68000,
      70,
      780,
      5.0,
      'https://images.unsplash.com/photo-1559056199-641a0ac8b55e?auto=format&fit=crop&w=600&q=80',
      'Bubuk Halus,Bubuk Sedang,Biji Utuh',
      false,
      'Mall',
      '100% Kopi Arabika Gayo asli dari dataran tinggi Aceh. Dipanggang medium roast dengan aroma rempah dan coklat yang khas.',
      'SUP-003',
      'UD Pangan Nusantara Mandiri'
    ]
  ];

  dataProduk.forEach(row => sheetProduk.appendRow(row));

  // 4. SETUP SHEET PESANAN
  let sheetPesanan = ss.getSheetByName('Pesanan');
  if (!sheetPesanan) {
    sheetPesanan = ss.insertSheet('Pesanan');
  } else {
    sheetPesanan.clear();
  }

  const headerPesanan = [
    'ID_Pesanan', 'Tanggal', 'Nama_Pembeli', 'No_WA', 'Alamat',
    'Catatan', 'Daftar_Item_JSON', 'Ringkasan_Item', 'Subtotal',
    'Ongkir', 'Diskon_Voucher', 'Total_Bayar', 'Kurir',
    'Metode_Bayar', 'Status', 'No_Resi'
  ];
  sheetPesanan.appendRow(headerPesanan);
  styleHeader(sheetPesanan, headerPesanan.length);

  const sampleItems = JSON.stringify([
    { id: 'PRD-001', nama: 'TWS Earbuds Bluetooth 5.3 Bass Hifi', varian: 'Hitam', harga: 99000, qty: 1 }
  ]);
  sheetPesanan.appendRow([
    'ORD-1001',
    Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss'),
    'Budi Santoso',
    '081234567890',
    'Jl. Sudirman No. 45, Kec. Setiabudi, Jakarta Selatan',
    'Tolong packing bubble wrap tebal ya kak',
    sampleItems,
    '1x TWS Earbuds Bluetooth 5.3 (Hitam)',
    99000,
    10000,
    0,
    109000,
    'Reguler',
    'QRIS',
    'Dikirim',
    'JP9981234567'
  ]);

  // 5. SETUP SHEET VOUCHER
  let sheetVoucher = ss.getSheetByName('Voucher');
  if (!sheetVoucher) {
    sheetVoucher = ss.insertSheet('Voucher');
  } else {
    sheetVoucher.clear();
  }

  const headerVoucher = [
    'Kode_Voucher', 'Tipe', 'Nilai_Potongan', 'Min_Belanja', 'Kuota', 'Deskripsi'
  ];
  sheetVoucher.appendRow(headerVoucher);
  styleHeader(sheetVoucher, headerVoucher.length);

  // 6. SETUP SHEET ULASAN
  let sheetUlasan = ss.getSheetByName('Ulasan');
  if (!sheetUlasan) {
    sheetUlasan = ss.insertSheet('Ulasan');
  } else {
    sheetUlasan.clear();
  }

  const headerUlasan = [
    'ID_Ulasan', 'ID_Produk', 'Nama_Pembeli', 'Rating', 'Komentar', 'Varian', 'Tanggal'
  ];
  sheetUlasan.appendRow(headerUlasan);
  styleHeader(sheetUlasan, headerUlasan.length);

  const dataUlasan = [
    ['REV-001', 'PRD-001', 'Andi P***', 5, 'Suara bass mantap banget untuk harga segini! Koneksi ke HP cepat.', 'Hitam', '10 Februari 2025, 14:20'],
    ['REV-002', 'PRD-001', 'Siti N***', 5, 'Warna pink pastelnya lucu, pengiriman juga cepat sampai.', 'Pink Pastel', '12 Februari 2025, 09:45'],
    ['REV-003', 'PRD-002', 'Rizky F***', 5, 'Kain flanelnya lembut dan tebal, ukuran L pas di badan saya.', 'Ukuran L', '14 Februari 2025, 19:10'],
    ['REV-004', 'PRD-003', 'Dewi K***', 5, 'Layar AMOLED-nya jernih banget, kelihatan seperti jam tangan mahal!', 'Rose Gold', '15 Februari 2025, 11:30'],
    ['REV-005', 'PRD-005', 'Fajar M***', 5, 'Coklat lumernya melimpah, pisangnya gak keras sama sekali.', 'Coklat Lumer', '16 Februari 2025, 16:05']
  ];
  dataUlasan.forEach(row => sheetUlasan.appendRow(row));

  // Auto-resize seluruh kolom agar rapi
  [sheetSupplier, sheetPengaturan, sheetProduk, sheetPesanan, sheetVoucher, sheetUlasan].forEach(sheet => {
    sheet.autoResizeColumns(1, sheet.getLastColumn());
  });

  Logger.log('Setup Database Hibualamo Store berhasil diselesaikan!');
}

/**
 * Fungsi pembantu untuk memberi warna header tabel (#EE4D2D)
 */
function styleHeader(sheet, colCount) {
  const range = sheet.getRange(1, 1, 1, colCount);
  range.setBackground('#ee4d2d');
  range.setFontColor('#ffffff');
  range.setFontWeight('bold');
  range.setHorizontalAlignment('center');
  sheet.setFrozenRows(1);
}