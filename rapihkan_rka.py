"""
================================================================================
SCRIPT AUTOMASI: PERAPI RKA BELANJA SIPD KE EXCEL COMPACT & RAPI
================================================================================
Fungsi:
1. Membaca file RKA hasil konversi PDFgear yang berantakan (banyak kolom/baris liar).
2. Mengekstrak 44 kode rekening beserta nominal belanja (Sebelum & Sesudah).
3. Mengatur lebar kolom Nomor Rekening secara presisi berbasis digit terpanjang (21 digit).
4. Menyusun tata letak hemat ruang (compact), menghilangkan merge cell liar dan baris ganda.
5. Memberikan indentasi hierarki akun (Level 1 s/d Level 6) serta pewarnaan visual.
6. Menerapkan rumus Excel dinamis untuk Selisih dan % Perubahan.
7. Menghasilkan output .xlsx modern dan .xls klasik.
================================================================================
"""

import os
import sys
import subprocess

def rapihkan_rka(source_path=None, output_dir=None):
    base_dir = r"C:\Users\a_ros\Documents\PDFgear"
    
    if not source_path:
        source_path = os.path.join(
            base_dir,
            "Sistem Informasi Pemerintahan Daerah - Cetak RKA Rincian Belanja _ 2.16.01.2.02.0001 Penyediaan Gaji dan Tunjangan ASN conv 1.xls"
        )
        
    if not output_dir:
        output_dir = os.path.dirname(source_path)

    output_xlsx = os.path.join(output_dir, "RKA_Penyediaan_Gaji_dan_Tunjangan_ASN_2026_COMPACT.xlsx")
    output_xls  = os.path.join(output_dir, "RKA_Penyediaan_Gaji_dan_Tunjangan_ASN_2026_COMPACT.xls")

    if not os.path.exists(source_path):
        print(f"[ERROR] File sumber tidak ditemukan di:\n{source_path}")
        return

    print("=" * 70)
    print("MEMULAI PROSES PERAPIAN FILE RKA EXCEL")
    print(f"File Sumber : {source_path}")
    print(f"Target XLSX : {output_xlsx}")
    print("=" * 70)

    # Engine automasi Excel COM Native Windows
    ps_script = f'''
$ErrorActionPreference = 'Stop'
$sourcePath = '{source_path}'
$destXlsx   = '{output_xlsx}'
$destXls    = '{output_xls}'

# Tutup file jika sudah ada sebelumnya agar tidak terkunci
if (Test-Path $destXlsx) {{ Remove-Item $destXlsx -Force }}
if (Test-Path $destXls)  {{ Remove-Item $destXls -Force }}

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

try {{
    # 1. Buka file sumber dan ekstrak data
    $wbSrc = $excel.Workbooks.Open($sourcePath)
    $shSrc = $wbSrc.Sheets.Item(1)
    
    $items = @()
    for ($r = 45; $r -le 172; $r++) {{
        $rek = $shSrc.Cells.Item($r, 2).Text.Trim()
        $uraian = $shSrc.Cells.Item($r, 4).Text.Trim()
        $sblm = $shSrc.Cells.Item($r, 15).Text.Trim()
        $ssdh = $shSrc.Cells.Item($r, 24).Text.Trim()
        
        if ($rek -match '^[0-9\\.]+$') {{
            $valSblm = [double]($sblm.Replace('.', '').Replace(',', '.'))
            $valSsdh = [double]($ssdh.Replace('.', '').Replace(',', '.'))
            
            $items += [PSCustomObject]@{{
                Kode = $rek
                Uraian = $uraian
                Sebelum = $valSblm
                Sesudah = $valSsdh
            }}
        }}
    }}
    $wbSrc.Close($false)
    Write-Host ("Berhasil mengekstrak " + $items.Count + " item rekening.")

    # 2. Buat Workbook Baru
    $wbNew = $excel.Workbooks.Add()
    while ($wbNew.Sheets.Count -gt 1) {{
        $wbNew.Sheets.Item(2).Delete()
    }}
    
    # -------------------------------------------------------------
    # SHEET 1: RKA_Kompak (Tampilan Ringkas & Hemat Ruang)
    # -------------------------------------------------------------
    $sh1 = $wbNew.Sheets.Item(1)
    $sh1.Name = "RKA_Kompak"
    $sh1.Activate()
    $excel.ActiveWindow.DisplayGridlines = $true
    
    $sh1.Cells.Font.Name = "Segoe UI"
    $sh1.Cells.Font.Size = 9.5
    
    # Kop Header RKA
    $sh1.Cells.Item(1, 1).Value2 = "PEMERINTAH KABUPATEN BLITAR"
    $sh1.Cells.Item(1, 1).Font.Size = 13
    $sh1.Cells.Item(1, 1).Font.Bold = $true
    
    $sh1.Cells.Item(2, 1).Value2 = "RENCANA KERJA DAN ANGGARAN (RKA-BELANJA PERUBAHAN) - TAHUN ANGGARAN 2026"
    $sh1.Cells.Item(2, 1).Font.Size = 10.5
    $sh1.Cells.Item(2, 1).Font.Bold = $true
    
    $sh1.Cells.Item(4, 1).Value2 = "SKPD"
    $sh1.Cells.Item(4, 2).Value2 = ": Dinas Komunikasi, Informatika, Statistik dan Persandian (2.16.2.20.2.21.01.0000)"
    $sh1.Cells.Item(5, 1).Value2 = "Sub Kegiatan"
    $sh1.Cells.Item(5, 2).Value2 = ": 2.16.01.2.02.0001 Penyediaan Gaji dan Tunjangan ASN"
    $sh1.Cells.Item(6, 1).Value2 = "Sumber Dana"
    $sh1.Cells.Item(6, 2).Value2 = ": Dana Alokasi Umum (DAU) | Target: 37 Orang / bulan"
    $sh1.Range("A4:A6").Font.Bold = $true
    
    # Header Tabel
    $tblStart = 8
    $headers = @(
        "Kode Rekening",
        "Uraian Belanja",
        "Sumber Dana",
        "Sebelum Perubahan (Rp)",
        "Sesudah Perubahan (Rp)",
        "Bertambah / (Berkurang) (Rp)",
        "% Perubahan"
    )
    
    for ($c = 0; $c -lt $headers.Count; $c++) {{
        $cell = $sh1.Cells.Item($tblStart, $c + 1)
        $cell.Value2 = $headers[$c]
        $cell.Font.Bold = $true
        $cell.Font.Color = 0xFFFFFF
        $cell.Interior.Color = 0x5D361F # Dark Slate Navy
        $cell.HorizontalAlignment = -4108 # Center
        $cell.VerticalAlignment = -4108
        $cell.RowHeight = 26
    }}
    
    $currRow = $tblStart + 1
    foreach ($it in $items) {{
        $dots = ($it.Kode.ToCharArray() | Where-Object {{ $_ -eq '.' }}).Count
        
        # Kolom 1: Kode Rekening (Acuan lebar: digit terpanjang)
        $cA = $sh1.Cells.Item($currRow, 1)
        $cA.NumberFormat = "@"
        $cA.Value2 = $it.Kode
        $cA.HorizontalAlignment = -4131 # Left
        
        # Kolom 2: Uraian Belanja + Indentasi Hirarki
        $cB = $sh1.Cells.Item($currRow, 2)
        $cB.Value2 = $it.Uraian
        $cB.IndentLevel = [Math]::Min($dots, 5)
        
        # Kolom 3: Sumber Dana
        $cC = $sh1.Cells.Item($currRow, 3)
        if ($dots -ge 4) {{ $cC.Value2 = "DAU" }}
        $cC.HorizontalAlignment = -4108
        
        # Kolom 4: Sebelum (Rp)
        $cD = $sh1.Cells.Item($currRow, 4)
        $cD.Formula = [string]$it.Sebelum
        $cD.NumberFormatLocal = '#.##0'
        $cD.HorizontalAlignment = -4152
        
        # Kolom 5: Sesudah (Rp)
        $cE = $sh1.Cells.Item($currRow, 5)
        $cE.Formula = [string]$it.Sesudah
        $cE.NumberFormatLocal = '#.##0'
        $cE.HorizontalAlignment = -4152
        
        # Kolom 6: Selisih (Formula Excel Dinamis)
        $cF = $sh1.Cells.Item($currRow, 6)
        $cF.Formula = ('=E' + $currRow + '-D' + $currRow)
        $cF.NumberFormatLocal = '#.##0;[Red]-#.##0;""-""'
        $cF.HorizontalAlignment = -4152
        
        # Kolom 7: % Perubahan (Formula Excel Dinamis)
        $cG = $sh1.Cells.Item($currRow, 7)
        $cG.Formula = ('=IF(D' + $currRow + '=0, 0, (E' + $currRow + '-D' + $currRow + ')/D' + $currRow + ')')
        $cG.NumberFormatLocal = '0,00%;[Red]-0,00%;0,00%'
        $cG.HorizontalAlignment = -4152
        
        # Pewarnaan & Border Baris
        $rowRng = $sh1.Range(('A' + $currRow + ':G' + $currRow))
        $rowRng.RowHeight = 20
        if ($dots -eq 0) {{
            $rowRng.Font.Bold = $true
            $rowRng.Interior.Color = 0xF0D5C7
        }} elseif ($dots -eq 1) {{
            $rowRng.Font.Bold = $true
            $rowRng.Interior.Color = 0xF0E0D0
        }} elseif ($dots -eq 2) {{
            $rowRng.Font.Bold = $true
            $rowRng.Interior.Color = 0xF5F5F5
        }} elseif ($dots -eq 3) {{
            $rowRng.Font.Bold = $true
        }} elseif ($dots -eq 4) {{
            $sh1.Range(('A' + $currRow + ':B' + $currRow)).Font.Bold = $true
        }}
        
        $rowRng.Borders.Item(11).LineStyle = 1
        $rowRng.Borders.Item(11).Color = 0xE0E0E0
        $rowRng.Borders.Item(9).LineStyle = 1
        $rowRng.Borders.Item(9).Color = 0xEDEDED
        
        $currRow++
    }}
    
    # Baris Total Rekapitulasi di bawah
    $totRow = $currRow
    $sh1.Cells.Item($totRow, 1).Value2 = "TOTAL BELANJA DAERAH (KODE REKENING 5)"
    $totMerged = $sh1.Range(('A' + $totRow + ':C' + $totRow))
    $totMerged.Merge()
    $totMerged.HorizontalAlignment = -4152
    $totMerged.Font.Bold = $true
    
    $sh1.Cells.Item($totRow, 4).Formula = "=D9"
    $sh1.Cells.Item($totRow, 4).NumberFormatLocal = '#.##0'
    $sh1.Cells.Item($totRow, 5).Formula = "=E9"
    $sh1.Cells.Item($totRow, 5).NumberFormatLocal = '#.##0'
    $sh1.Cells.Item($totRow, 6).Formula = "=F9"
    $sh1.Cells.Item($totRow, 6).NumberFormatLocal = '#.##0;[Red]-#.##0;""-""'
    $sh1.Cells.Item($totRow, 7).Formula = "=G9"
    $sh1.Cells.Item($totRow, 7).NumberFormatLocal = '0,00%;[Red]-0,00%;0,00%'
    
    $totRng = $sh1.Range(('A' + $totRow + ':G' + $totRow))
    $totRng.Font.Bold = $true
    $totRng.RowHeight = 24
    $totRng.Interior.Color = 0xE8D0C0
    $totRng.Borders.Item(8).LineStyle = 1
    $totRng.Borders.Item(9).LineStyle = -4119
    
    # Border luar
    $tblRng = $sh1.Range(('A' + $tblStart + ':G' + $totRow))
    $tblRng.Borders.Item(7).LineStyle = 1
    $tblRng.Borders.Item(10).LineStyle = 1
    $tblRng.Borders.Item(8).LineStyle = 1
    
    # PENGATURAN LEBAR KOLOM (ACUAN: DIGIT TERPANJANG 21 KARAKTER)
    $sh1.Columns.Item(1).ColumnWidth = 24.0 # Kode Rekening pas dengan '5.1.01.01.001.00001'
    $sh1.Columns.Item(2).ColumnWidth = 50.0 # Uraian
    $sh1.Columns.Item(3).ColumnWidth = 14.0 # Sumber Dana
    $sh1.Columns.Item(4).ColumnWidth = 22.0 # Sebelum
    $sh1.Columns.Item(5).ColumnWidth = 22.0 # Sesudah
    $sh1.Columns.Item(6).ColumnWidth = 24.0 # Selisih
    $sh1.Columns.Item(7).ColumnWidth = 14.0 # %
    
    # Aktifkan AutoFilter & Freeze Panes
    $sh1.Range(('A' + $tblStart + ':G' + $tblStart)).AutoFilter() | Out-Null
    $sh1.Cells.Item(9, 1).Select() | Out-Null
    $excel.ActiveWindow.FreezePanes = $true

    # -------------------------------------------------------------
    # SHEET 2: RKA_Rincian_Standar (Format Lengkap SIPD)
    # -------------------------------------------------------------
    $sh2 = $wbNew.Sheets.Add([System.Reflection.Missing]::Value, $sh1)
    $sh2.Name = "RKA_Rincian_Standar"
    $sh2.Activate()
    $excel.ActiveWindow.DisplayGridlines = $true
    $sh2.Cells.Font.Name = "Segoe UI"
    $sh2.Cells.Font.Size = 9.5
    
    $sh2.Cells.Item(1, 1).Value2 = "RINCIAN ANGGARAN BELANJA SUB KEGIATAN (FORMAT STANDAR LENGKAP SIPD)"
    $sh2.Cells.Item(1, 1).Font.Size = 12
    $sh2.Cells.Item(1, 1).Font.Bold = $true
    
    $sh2.Range("A3:A4").Merge()
    $sh2.Cells.Item(3, 1).Value2 = "Kode Rekening"
    $sh2.Range("B3:B4").Merge()
    $sh2.Cells.Item(3, 2).Value2 = "Uraian"
    
    $sh2.Range("C3:G3").Merge()
    $sh2.Cells.Item(3, 3).Value2 = "Rincian Perhitungan Sebelum"
    $sh2.Cells.Item(4, 3).Value2 = "Koefisien"
    $sh2.Cells.Item(4, 4).Value2 = "Satuan"
    $sh2.Cells.Item(4, 5).Value2 = "Harga (Rp)"
    $sh2.Cells.Item(4, 6).Value2 = "PPN"
    $sh2.Cells.Item(4, 7).Value2 = "Jumlah (Rp)"
    
    $sh2.Range("H3:L3").Merge()
    $sh2.Cells.Item(3, 8).Value2 = "Rincian Perhitungan Sesudah"
    $sh2.Cells.Item(4, 8).Value2 = "Koefisien"
    $sh2.Cells.Item(4, 9).Value2 = "Satuan"
    $sh2.Cells.Item(4, 10).Value2 = "Harga (Rp)"
    $sh2.Cells.Item(4, 11).Value2 = "PPN"
    $sh2.Cells.Item(4, 12).Value2 = "Jumlah (Rp)"
    
    $sh2.Range("M3:M4").Merge()
    $sh2.Cells.Item(3, 13).Value2 = "Bertambah / (Berkurang) (Rp)"
    
    $hdrRng2 = $sh2.Range("A3:M4")
    $hdrRng2.Font.Bold = $true
    $hdrRng2.Font.Color = 0xFFFFFF
    $hdrRng2.Interior.Color = 0x5D361F
    $hdrRng2.HorizontalAlignment = -4108
    $hdrRng2.VerticalAlignment = -4108
    
    $r2 = 5
    foreach ($it in $items) {{
        $dots = ($it.Kode.ToCharArray() | Where-Object {{ $_ -eq '.' }}).Count
        
        $sh2.Cells.Item($r2, 1).NumberFormat = "@"
        $sh2.Cells.Item($r2, 1).Value2 = $it.Kode
        $sh2.Cells.Item($r2, 2).Value2 = $it.Uraian
        $sh2.Cells.Item($r2, 2).IndentLevel = [Math]::Min($dots, 5)
        
        if ($dots -ge 5) {{
            $sh2.Cells.Item($r2, 3).Value2 = "1 Tahun"
            $sh2.Cells.Item($r2, 4).Value2 = "Tahun"
            $sh2.Cells.Item($r2, 5).Formula = [string]$it.Sebelum
            $sh2.Cells.Item($r2, 5).NumberFormatLocal = '#.##0'
            $sh2.Cells.Item($r2, 6).Value2 = "-"
            $sh2.Cells.Item($r2, 6).HorizontalAlignment = -4108
            
            $sh2.Cells.Item($r2, 8).Value2 = "1 Tahun"
            $sh2.Cells.Item($r2, 9).Value2 = "Tahun"
            $sh2.Cells.Item($r2, 10).Formula = [string]$it.Sesudah
            $sh2.Cells.Item($r2, 10).NumberFormatLocal = '#.##0'
            $sh2.Cells.Item($r2, 11).Value2 = "-"
            $sh2.Cells.Item($r2, 11).HorizontalAlignment = -4108
        }}
        
        $sh2.Cells.Item($r2, 7).Formula = [string]$it.Sebelum
        $sh2.Cells.Item($r2, 7).NumberFormatLocal = '#.##0'
        
        $sh2.Cells.Item($r2, 12).Formula = [string]$it.Sesudah
        $sh2.Cells.Item($r2, 12).NumberFormatLocal = '#.##0'
        
        $sh2.Cells.Item($r2, 13).Formula = ('=L' + $r2 + '-G' + $r2)
        $sh2.Cells.Item($r2, 13).NumberFormatLocal = '#.##0;[Red]-#.##0;""-""'
        
        $rowRng2 = $sh2.Range(('A' + $r2 + ':M' + $r2))
        $rowRng2.RowHeight = 20
        if ($dots -eq 0) {{
            $rowRng2.Font.Bold = $true
            $rowRng2.Interior.Color = 0xF0D5C7
        }} elseif ($dots -le 2) {{
            $rowRng2.Font.Bold = $true
            $rowRng2.Interior.Color = 0xF5F5F5
        }} elseif ($dots -le 4) {{
            $sh2.Range(('A' + $r2 + ':B' + $r2)).Font.Bold = $true
        }}
        
        $rowRng2.Borders.Item(11).LineStyle = 1
        $rowRng2.Borders.Item(11).Color = 0xE0E0E0
        $rowRng2.Borders.Item(9).LineStyle = 1
        $rowRng2.Borders.Item(9).Color = 0xEDEDED
        
        $r2++
    }}
    
    $sh2.Columns.Item(1).ColumnWidth = 24.0 # Kode Rekening (ACUAN DIGIT TERPANJANG)
    $sh2.Columns.Item(2).ColumnWidth = 46.0 # Uraian
    $sh2.Columns.Item(3).ColumnWidth = 11.0 # Koef Sblm
    $sh2.Columns.Item(4).ColumnWidth = 9.0  # Satuan Sblm
    $sh2.Columns.Item(5).ColumnWidth = 18.0 # Harga Sblm
    $sh2.Columns.Item(6).ColumnWidth = 7.0  # PPN Sblm
    $sh2.Columns.Item(7).ColumnWidth = 19.0 # Jumlah Sblm
    $sh2.Columns.Item(8).ColumnWidth = 11.0 # Koef Ssdh
    $sh2.Columns.Item(9).ColumnWidth = 9.0  # Satuan Ssdh
    $sh2.Columns.Item(10).ColumnWidth = 18.0# Harga Ssdh
    $sh2.Columns.Item(11).ColumnWidth = 7.0 # PPN Ssdh
    $sh2.Columns.Item(12).ColumnWidth = 19.0# Jumlah Ssdh
    $sh2.Columns.Item(13).ColumnWidth = 22.0# Selisih
    
    $sh2.Cells.Item(5, 1).Select() | Out-Null
    $excel.ActiveWindow.FreezePanes = $true

    # -------------------------------------------------------------
    # SHEET 3: Info_Sub_Kegiatan (Metadata & Indikator Kinerja)
    # -------------------------------------------------------------
    $sh3 = $wbNew.Sheets.Add([System.Reflection.Missing]::Value, $sh2)
    $sh3.Name = "Info_Sub_Kegiatan"
    $sh3.Activate()
    $excel.ActiveWindow.DisplayGridlines = $true
    $sh3.Cells.Font.Name = "Segoe UI"
    $sh3.Cells.Font.Size = 10
    
    $sh3.Cells.Item(1, 1).Value2 = "INFORMASI KEGIATAN & INDIKATOR TOLAK UKUR KINERJA"
    $sh3.Cells.Item(1, 1).Font.Size = 12
    $sh3.Cells.Item(1, 1).Font.Bold = $true
    
    $info = @(
        @("Urusan Pemerintahan", "2 URUSAN PEMERINTAHAN WAJIB YANG TIDAK BERKAITAN DENGAN PELAYANAN DASAR"),
        @("Bidang Urusan", "2.16 URUSAN PEMERINTAHAN BIDANG KOMUNIKASI DAN INFORMATIKA"),
        @("Unit Organisasi", "2.16.2.20.2.21.01.0000 Dinas Komunikasi, Informatika, Statistik dan Persandian"),
        @("Sub Unit Organisasi", "2.16.2.20.2.21.01.0000 Dinas Komunikasi, Informatika, Statistik dan Persandian"),
        @("Program", "2.16.01 PROGRAM PENUNJANG URUSAN PEMERINTAHAN DAERAH KABUPATEN/KOTA"),
        @("Kegiatan", "2.16.01.2.02 Administrasi Keuangan Perangkat Daerah"),
        @("Sub Kegiatan", "2.16.01.2.02.0001 Penyediaan Gaji dan Tunjangan ASN"),
        @("Sumber Pendanaan", "Dana Alokasi Umum (DAU)"),
        @("Lokasi Pelaksanaan", "Semua Kota/Kab, Semua Kecamatan, Semua Kel/Desa"),
        @("Waktu Pelaksanaan", "Januari s.d Desember"),
        @("Alokasi Tahun 2025", "Rp 0,00"),
        @("Alokasi Tahun 2026", "Rp 3.674.068.000,00"),
        @("Alokasi Tahun 2027", "Rp 0,00"),
        @("Indikator Masukan (Sebelum)", "Dana yang dibutuhkan : Rp 3.985.562.324,00"),
        @("Indikator Masukan (Sesudah)", "Dana yang dibutuhkan : Rp 3.674.068.000,00"),
        @("Indikator Keluaran", "Jumlah Orang yang Menerima Gaji dan Tunjangan ASN : 37 Orang/bulan")
    )
    
    for ($i = 0; $i -lt $info.Count; $i++) {{
        $rowIdx = $i + 3
        $sh3.Cells.Item($rowIdx, 1).Value2 = $info[$i][0]
        $sh3.Cells.Item($rowIdx, 1).Font.Bold = $true
        $sh3.Cells.Item($rowIdx, 2).Value2 = $info[$i][1]
        
        $sh3.Range(('A' + $rowIdx + ':B' + $rowIdx)).Borders.Item(9).LineStyle = 1
        $sh3.Range(('A' + $rowIdx + ':B' + $rowIdx)).Borders.Item(9).Color = 0xE0E0E0
        $sh3.Rows.Item($rowIdx).RowHeight = 22
    }}
    $sh3.Columns.Item(1).ColumnWidth = 28.0
    $sh3.Columns.Item(2).ColumnWidth = 75.0
    
    # Kembali aktifkan Sheet 1 sebagai tampilan utama
    $sh1.Activate()
    $sh1.Cells.Item(9, 1).Select() | Out-Null
    
    # Simpan file
    Write-Host "Menyimpan format XLSX..."
    $wbNew.SaveAs($destXlsx, 51)
    
    Write-Host "Menyimpan format XLS..."
    $wbNew.SaveAs($destXls, 56)
    
    $wbNew.Close($false)
    Write-Host "SELESAI DENGAN SUKSES!"
}} finally {{
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}}
'''

    temp_ps1 = os.path.join(output_dir, "_temp_engine.ps1")
    try:
        with open(temp_ps1, "w", encoding="utf-8") as f:
            f.write(ps_script)

        cmd = ["powershell", "-ExecutionPolicy", "Bypass", "-File", temp_ps1]
        res = subprocess.run(cmd, capture_output=True, text=True)
        
        if res.returncode == 0:
            print(res.stdout)
            print("=" * 70)
            print("[BERHASIL] File Excel Rapi & Compact berhasil dibuat!")
            print(f"-> {output_xlsx}")
            print(f"-> {output_xls}")
            print("=" * 70)
        else:
            print("[GAGAL] Terjadi kesalahan saat eksekusi:")
            print(res.stderr)
    finally:
        if os.path.exists(temp_ps1):
            os.remove(temp_ps1)

if __name__ == "__main__":
    rapihkan_rka()
