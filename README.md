# Generator Transkrip Nilai — GitHub Pages

Ini antarmuka web Generator Transkrip Nilai SMAIT Al-Ittihad yang di-hosting di **GitHub Pages**. Semua data dan proses (Google Sheets, Google Docs, PDF/DOCX) tetap berjalan di **Google Apps Script**:

```
https://script.google.com/macros/s/AKfycbx3-l2dVp_GjlZNiFwYF6ti7gjlLgdpm5R1KnWE4e_AEKLAzt0YrZ52jbQtKBaTtIwK/exec
```

```
Browser (GitHub Pages)  ──POST JSON + kode akses──▶  Apps Script /exec  ──▶  Sheets · Docs · Drive
        index.html · app.js · config.js           ◀──────── { ok, data } ────────
```

## Isi folder

| File | Fungsi |
|---|---|
| `index.html` | Halaman aplikasi (sama dengan versi Apps Script, ditambah layar masuk dan tombol Keluar) |
| `app.js` | Logika frontend. Pemanggilan `google.script.run` diganti dengan `fetch` ke URL `/exec` |
| `config.js` | **URL Web App Apps Script.** Ganti di sini bila URL berubah |
| `.nojekyll` | Agar GitHub Pages menyajikan file apa adanya |

## Langkah 1 — Perbarui Apps Script (wajib, sekali saja)

1. Buka proyek Apps Script, lalu ganti isi **`Code.gs`** dengan versi 2.2.0. File `Index.html` dan `JavaScript.html` tidak berubah.
2. Buka **Project Settings → "Show appsscript.json"**, lalu samakan `webapp` dengan:
   ```json
   "webapp": { "executeAs": "USER_DEPLOYING", "access": "ANYONE_ANONYMOUS" }
   ```
3. Jalankan fungsi **`setupDatabase`** sekali. Fungsi ini membuat **kode akses** (`TRX-…`) di sheet `Pengaturan` (kunci `API_TOKEN`) dan menampilkannya.
   Kode itu bisa dilihat lagi kapan saja lewat menu spreadsheet **🎓 Transkrip → Tampilkan kode akses (GitHub Pages)**.
4. Deploy ulang **tanpa mengubah URL**: **Deploy → Manage deployments → ✏️ Edit**, lalu atur:
   - **Version:** *New version*
   - **Execute as:** *Me*
   - **Who has access:** **Anyone** (Siapa saja). Halaman GitHub memanggil API tanpa login Google, jadi pengaturan ini wajib.

   Klik **Deploy**. URL `/exec` tetap sama dengan yang ada di `config.js`.

> **Apakah aman memilih "Anyone"?** Ya. Semua aksi API mewajibkan kode akses. Tanpa kode yang benar, server hanya menjawab `Kode akses salah`. Antarmuka bawaan di URL `/exec` juga tertutup untuk publik; pemilik bisa membukanya lewat `…/exec?ui=1&token=KODE_AKSES`. Jika kode bocor, jalankan menu **🎓 Transkrip → Buat kode akses baru**. Kode lama langsung tidak berlaku.

## Langkah 2 — Unggah ke GitHub

1. Buat repository baru di GitHub, misalnya `transkrip-nilai`. Visibilitasnya boleh *Public*, karena kode akses tidak tersimpan di file mana pun.
2. Unggah **isi** folder ini (`index.html`, `app.js`, `config.js`, `.nojekyll`, `README.md`) ke root repository. Caranya: **Add file → Upload files**, lalu *Commit*.
   - Lewat Git:
     ```bash
     git init && git add . && git commit -m "Generator Transkrip Nilai"
     git branch -M main
     git remote add origin https://github.com/USERNAME/transkrip-nilai.git
     git push -u origin main
     ```
3. Buka **Settings → Pages**. Pada *Build and deployment*, pilih **Source: Deploy from a branch**, lalu **Branch: `main` / `(root)`**, kemudian **Save**.
4. Tunggu 1–2 menit. Situs akan aktif di `https://USERNAME.github.io/transkrip-nilai/`.

## Langkah 3 — Masuk

Buka alamat GitHub Pages, lalu masukkan **kode akses** (`TRX-…`).
- Jika **"Ingat di perangkat ini"** dicentang, kode tersimpan di browser tersebut. Jangan centang di komputer bersama.
- Tombol **Keluar** di pojok kanan atas menghapus kode dari browser.

## Pemecahan masalah

| Pesan | Penyebab dan solusi |
|---|---|
| *Server meminta login Google…* | Deployment masih "Only myself" atau "Anyone with Google account". Ubah ke **Anyone**, lalu deploy versi baru. |
| *Kode akses salah* | Salin ulang dari menu **🎓 Transkrip → Tampilkan kode akses**. Mungkin juga kode baru saja diganti. |
| *API_TOKEN belum diatur…* | Jalankan `setupDatabase` atau `buatKodeAksesBaru` di editor Apps Script. |
| *Aksi tidak dikenal: …* | `Code.gs` di Apps Script masih versi lama, atau belum di-deploy sebagai **New version**. |
| *Tidak dapat terhubung…* | Periksa internet dan URL di `config.js` (harus berakhiran `/exec`, bukan `/dev`). |
| Perubahan di Apps Script tidak terlihat | Setiap perubahan `Code.gs` perlu **Manage deployments → Edit → New version**. |

Jika Anda memakai **"New deployment"** (bukan Edit), URL `/exec` berubah. Perbarui `GAS_URL` di `config.js`, lalu commit ulang.

## Untuk pengembang

`index.html` dan `app.js` dibangun otomatis dari `Index.html` dan `JavaScript.html` versi Apps Script:

```bash
python3 tools/build_github_pages.py
```

Jadi setiap perubahan cukup dibuat di file Apps Script, lalu jalankan skrip tersebut. `config.js` yang sudah ada tidak akan ditimpa.

API yang tersedia (POST body `text/plain` berisi JSON `{action, args, token}`):
`apiBootstrap, apiListStudents, apiGetStudent, apiSaveStudent, apiDeleteStudent, apiListSubjects, apiSaveSubject, apiDeleteSubject, apiImport, apiImportKurikulum, apiListJurusan, apiGetTranscriptPreview, apiGenerate, apiListTranscripts, apiRefreshFromDoc, apiDeleteTranscript, apiGetEditor, apiSaveEditor, apiDownload`, ditambah `ping` (tanpa kode akses).
