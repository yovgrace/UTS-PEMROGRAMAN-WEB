𝐔𝐓𝐒 𝐏𝐄𝐌𝐑𝐎𝐆𝐑𝐀𝐌𝐀𝐍 𝐖𝐄𝐁

𝐀𝐧𝐠𝐠𝐨𝐭𝐚 𝐊𝐞𝐥𝐨𝐦𝐩𝐨𝐤 :

𝟏. 𝐌𝐮𝐡𝐚𝐦𝐦𝐚𝐝 𝐍𝐚𝐛𝐢𝐥 𝐀𝐛𝐝𝐢𝐥𝐥𝐚𝐡_𝟐𝟎𝟐𝟒𝟏𝟐𝟎𝟓𝟑

𝟐. 𝐌𝐮𝐡𝐚𝐦𝐦𝐚𝐝 𝐃𝐰𝐢 𝐅𝐢𝐫𝐦𝐚𝐧𝐬𝐲𝐚𝐡_𝟐𝟎𝟐𝟒𝟏𝟐𝟎𝟏𝟐

𝟑. 𝐘𝐨𝐯𝐢𝐭𝐡𝐚 𝐆𝐫𝐚𝐜𝐢𝐚 𝐓𝐚𝐯𝐚𝐫𝐞𝐬_𝟐𝟎𝟐𝟒𝟏𝟐𝟎𝟒𝟒


𝐅𝐢𝐭𝐮𝐫 𝐭𝐚𝐦𝐛𝐚𝐡𝐚𝐧 :

𝟏. 𝐄𝐤𝐬𝐩𝐨𝐫 𝐄𝐱𝐜𝐞𝐥

<img width="1532" height="642" alt="Ekspor Excel" src="https://github.com/user-attachments/assets/2703f1fb-be88-4b53-87f2-8f5ec4b2d8ed" />


jadi intinya fungsi eksporExcel() ini gunanya buat nge-convert data mahasiswa di web jadi file .xlsx pake library SheetJS, di mana sistem bakal ngecek dulu ketersediaan datanya lewat if(data.length === 0) biar gak error, lalu ngelakuin mapping data pake .map() biar struktur kolomnya rapi (termasuk ngubah inisial 'L' atau 'P' jadi teks 'Laki-laki'/'Perempuan'), baru setelah itu datanya diubah jadi lembaran worksheet, dimasukin ke workbook baru, dan otomatis di-download ke device user dengan nama file Data_Mahasiswa.xlsx. Sementara itu, kode di bawahnya cuma buat nge-handle fitur live search biar tiap kali user ngetik halaman langsung balik ke halaman pertama, sama ada fungsi updateStatistik() yang otomatis jalan pas web pertama kali dimuat buat memperbarui data tampilan dashboard lo.

𝟐. 𝐋𝐚𝐛𝐞𝐥 𝐒𝐭𝐚𝐭𝐢𝐬𝐭𝐢𝐤

<img width="1532" height="642" alt="statistik" src="https://github.com/user-attachments/assets/941c605f-16bc-4a13-aea0-9cb2d4c532e5" />


stats-grid (Bungkus Utama Kartu Statistik)
display: grid;: Mengaktifkan mode CSS Grid. Ini dipakai biar kita bisa ngatur tata letak elemen di dalamnya jadi bentuk baris dan kolom dengan gampang.
grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));: Kunci rahasia bikin layout-nya responsif otomatis tanpa perlu media query. Artinya, browser bakal otomatis nambahin atau ngurangin jumlah kolom sesuai lebar layar. Batas minimal lebar kartunya $200\text{px}$, dan kalau ada sisa ruang bakal dibagi rata (1fr).
gap: 16px;: Ngasih jarak selebar $16\text{px}$ antar kartu statistik biar gak saling nempel.
margin-bottom: 25px;: Ngasih jarak space kosong ke bawah sebesar $25\text{px}$ biar gak mepet sama elemen di bawahnya (kayak tabel atau pagination).

stats-card (Desain Kotak/Kartu Satuan)
background: #0f172a;: Ngasih warna latar belakang kartu pake warna dark mode (biru dongker gelap/navy).
padding: 20px;: Ngasih ruang/jarak di dalam kartu sebesar $20\text{px}$ biar konten di dalemnya gak mepet ke pinggir garis kotak.
border-radius: 12px;: Bikin sudut-sudut kotak kartunya jadi melengkung (rounded) sebesar $12\text{px}$ biar kelihatan modern dan clean.
border: 1px solid #334155;: Ngasih garis tepi tipis warna abu-abu gelap biar kartunya kelihatan tegas dan punya dimensi.
text-align: center;: Bikin semua teks di dalam kartu posisinya otomatis rata tengah.

stats-card h3 (Styling Judul/Label Statistik)
font-size: 14px;: Ngatur ukuran font tulisan judul/label jadi agak kecil ($14\text{px}$).
color: #94a3b8;: Ngasih warna teks abu-abu muda agak pudar, tipikal warna teks pendukung di dark mode.
margin-bottom: 8px;: Ngasih jarak $8\text{px}$ di bawah judul biar gak terlalu nempel sama angka statistiknya.
text-transform: uppercase;: Bikin semua huruf di judulnya otomatis berubah jadi HURUF KAPITAL semua.
letter-spacing: 0.5px;: Ngasih renggangan tipis antar huruf sebesar $0.5\text{px}$ biar tulisan kapitalnya lebih enak dibaca.

stats-card p (Styling Angka/Data Statistik)
font-size: 28px;: Bikin ukuran angka statistiknya gede ($28\text{px}$) biar mencolok sebagai informasi utama.
font-weight: 700;: Bikin teks angka jadi tebal (bold).
color: #f8fafc;: Ngasih warna teks putih cerah biar kontras banget sama background kartu yang gelap, jadinya gampang dibaca.

pagination (Bungkus Tombol Halaman)
display: flex;: Mengaktifkan mode CSS Flexbox biar tombol-tombol angka di dalamnya bisa berjejer rapi secara horizontal (ke samping).
justify-content: center;: Memastikan posisi seluruh deretan tombol pagination berada tepat di tengah-tengah halaman secara horizontal.
gap: 6px;: Ngasih jarak tipis sebesar $6\text{px}$ antar tombol angka biar gak dempetan pas mau diklik.
margin-top: 20px;: Ngasih jarak kosong ke atas sebesar $20\text{px}$ biar kontainer pagination ini gak nempel sama bagian tabel atau konten di atasnya.
