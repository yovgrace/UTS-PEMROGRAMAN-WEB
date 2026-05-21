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
