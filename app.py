import streamlit as st
import pandas as pd
import numpy as np
import io

st.title('Aplikasi Pengolahan THC Simpanan')
st.markdown("""
## File yang dibutuhkan
1. **DNR.xlsx**
   - Tarik data yang ada di https://drive.google.com/drive/folders/1WBKDd_XUfZ-qJ9uJ6z1qUeDjaynZY5IO.
   - Filter cabang yang akan di cek.
   - Buat File Baru dengan nama DNR.xlsx
   - Pisahkan per sheet antara anggota dan suami meninggal.
   - Agar meminimalisir error buat nama sheet nya "Anggota" dan "Suami" pada excelnya.

2. **Data Anggota.xlsx**
   - Tarik data dari MDIS di menu Laporan Ops Cabang → Detail Nasabah SRSS.
   - Kolom yang di perlukan :
    | No | Cabang | Center | Kelompok | ID Anggota | Nama Anggota | Nama Sesuai KTP | Nama Suami | Alamat | Tgl. Gabung | NO. KTP |.
   - Nama sheet tidak usah diubah biarkan tetap **MdClientInfo**
""")

## FUNGSI FORMAT NOMOR
def format_no(no):
    try:
        if pd.notna(no):
            return f'{int(no):02d}.'
        else:
            return ''
    except (ValueError, TypeError):
        return str(no)

def format_center(center):
    try:
        if pd.notna(center):
            return f'{int(center):03d}'
        else:
            return ''
    except (ValueError, TypeError):
        return str(center)

def format_kelompok(kelompok):
    try:
        if pd.notna(kelompok):
            return f'{int(kelompok):02d}'
        else:
            return ''
    except (ValueError, TypeError):
        return str(kelompok)

#-------------------------- UPLOAD FILE --------------------------#
uploaded_files = st.file_uploader("Unggah file Excel", accept_multiple_files=True, type=["xlsx"])

if uploaded_files:
    dfs = {}
    for file in uploaded_files:
        try:
            excel_file = pd.ExcelFile(file, engine='openpyxl')
            
            for sheet_name in excel_file.sheet_names:
                df = pd.read_excel(excel_file, sheet_name=sheet_name)
                key = f"{file.name}_{sheet_name}"
                dfs[key] = df
            
            st.success(f"File {file.name} berhasil diunggah dan diproses.")
        except Exception as e:
            st.error(f"Terjadi kesalahan saat memproses file {file.name}: {str(e)}")

    # Pengkategorian dataframe
    if 'DNR.xlsx_Anggota' in dfs:
        df_dnr_anggota = dfs['DNR.xlsx_Anggota']
        st.success("Data DNR Anggota berhasil dimuat.")
    else:
        st.error("File DNR.xlsx tidak ditemukan atau tidak memiliki sheet 'Anggota'.")
        
    if 'DNR.xlsx_Suami' in dfs:
        df_dnr_suami = dfs['DNR.xlsx_Suami']
        st.success("Data DNR Suami berhasil dimuat.")
    else:
        st.warning("File DNR.xlsx tidak memiliki sheet 'Suami' atau belum diunggah.")
        
    if 'Data Anggota.xlsx_MdClientInfo' in dfs:
        df_data_anggota = dfs['Data Anggota.xlsx_MdClientInfo']
        if 'NO. KTP' in df_data_anggota.columns:
            df_data_anggota['NO. KTP'] = "'" + df_data_anggota['NO. KTP'].astype(str)
        st.success("Data Anggota berhasil dimuat.")
    else:
        st.error("File Data Anggota.xlsx tidak ditemukan atau tidak memiliki sheet 'MdClientInfo'.")

    # Lanjutkan hanya jika semua data yang diperlukan tersedia
    if all(key in dfs for key in ['DNR.xlsx_Anggota', 'Data Anggota.xlsx_MdClientInfo']):
        try:
            #----------------------DNR Anggota 
            # Tambah Kolom
            df_dnr_anggota['JENIS'] = 'ANGGOTA'

            # Ubah Nama Kolom
            rename_dict = {
                'No KTP': 'NO. KTP' 
            }
            df_dnr_anggota = df_dnr_anggota.rename(columns=rename_dict)

            # Merge DNR Anggota + Data Anggota
            merge_column = 'NO. KTP'
            df_agt_merge = pd.merge(df_dnr_anggota, df_data_anggota, on=merge_column, suffixes=('_df_agt','_df_data_agt'))

            # Ubah Nama Kolom Lagi
            rename_dict = {
                'NO. KTP_df_agt' : 'NO. KTP',
                'STATUS' : 'STATUS MENINGGAL',
                'TanggalPencairan' : 'TANGGAL CAIR',
                'Pokok' : 'DISBURSE',
                'Tanggal Kematian' : 'TANGGAL KEMATIAN',
                'PinjamanKe' : 'PINJ. KE-',
                'TanggalAprove DNR' : 'TANGGAL ACC DNR',
            }

            df_agt_merge = df_agt_merge.rename(columns=rename_dict)

            desired_order = [
                'No', 'NO. KTP', 'ID Anggota', 'Nama Anggota', 'Center', 'Kelompok', 'Nama Suami', 'Alamat', 'Tgl. Gabung', 'STATUS MENINGGAL', 'TANGGAL CAIR', 'DISBURSE','PINJ. KE-', 'TANGGAL KEMATIAN', 'TANGGAL ACC DNR'
            ]

            final_agt = df_agt_merge[desired_order]

            st.write("Anggota Meninggal:")
            st.write(final_agt)

            #----------------------DNR Suami 
            if 'DNR.xlsx_Suami' in dfs:
                # Tambah Kolom
                df_dnr_suami['JENIS'] = 'SUAMI' 

                #Ubah Nama Kolom
                rename_dict = {
                    'No KTP': 'NO. KTP' 
                }
                df_dnr_suami = df_dnr_suami.rename(columns=rename_dict)

                # Merge DNR Suami + Data Anggota
                merge_column = 'NO. KTP'
                df_suami_merge = pd.merge(df_dnr_suami, df_data_anggota, on=merge_column, suffixes=('_df_suami','_df_data_anggota'))

                # Ubah Nama Kolom Lagi
                rename_dict = {
                    'NO. KTP_df_suami' : 'NO. KTP',
                    'STATUS' : 'STATUS MENINGGAL',
                    'TanggalPencairan' : 'TANGGAL CAIR',
                    'Pokok' : 'DISBURSE',
                    'Tanggal Kematian' : 'TANGGAL KEMATIAN',
                    'PinjamanKe' : 'PINJ. KE-',
                    'TanggalAprove DNR' : 'TANGGAL ACC DNR',
                }

                df_suami_merge = df_suami_merge.rename(columns=rename_dict)

                desired_order = [
                    'No', 'NO. KTP', 'ID Anggota', 'Nama Anggota', 'Center', 'Kelompok', 'Nama Suami', 'Alamat', 'Tgl. Gabung', 'STATUS MENINGGAL', 'TANGGAL CAIR', 'DISBURSE','PINJ. KE-', 'TANGGAL KEMATIAN', 'TANGGAL ACC DNR'
                ]

                final_suami = df_suami_merge[desired_order]

                st.write("Suami Anggota Meninggal:")
                st.write(final_suami)
            else:
                st.info("Data DNR Suami tidak tersedia. Hanya menampilkan data Anggota Meninggal.")

        except Exception as e:
            st.error(f"Terjadi kesalahan saat memproses data: {str(e)}")
    else:
        st.warning("Mohon unggah semua file yang diperlukan (DNR.xlsx dan Data Anggota.xlsx) dengan format yang benar.")
else:
    st.info("Silakan unggah file Excel yang diperlukan.")
