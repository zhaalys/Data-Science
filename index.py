import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
import seaborn as sns

# Set style untuk grafik yang lebih menarik
try:
    plt.style.use('seaborn-v0_8')
except:
    plt.style.use('seaborn')
sns.set_palette("husl")

def nilaiSiswa():
    try:
        # Membaca data dari file excel
        data = pd.read_excel('Data_Siswa.xlsx')
        print("=== TAMPIL DATA SISWA ===")
        print(data)
        print("\n" + "="*50)

        # Menghitung rata-rata tiap siswa
        data["Rata_rata"] = data[["Matematika", "Bahasa Indonesia", "Informatika"]].mean(axis=1)
        data["Rata_rata"] = data["Rata_rata"].round(2)  # Bulatkan ke 2 desimal
        
        print("\n=== DATA DENGAN RATA-RATA ===")
        print(data)
        print("\n" + "="*50)

        # Menentukan siswa dengan nilai rata-rata tertinggi
        tertinggi = data.loc[data["Rata_rata"].idxmax()]
        print(f"\n=== SISWA DENGAN NILAI TERTINGGI ===")
        print(f"Nama: {tertinggi['Nama']}")
        print(f"Rata-rata: {tertinggi['Rata_rata']}")
        print("\n" + "="*50)
        
        return data
        
    except FileNotFoundError:
        print("Error: File 'Data_Siswa.xlsx' tidak ditemukan!")
        return None
    except Exception as e:
        print(f"Error: {str(e)}")
        return None

def buatVisualisasi(data):
    if data is None:
        print("Tidak dapat membuat visualisasi karena data tidak tersedia.")
        return
    
    # Membuat figure dengan multiple subplots
    fig, axes = plt.subplots(2, 2, figsize=(15, 12))
    fig.suptitle('ANALISIS NILAI SISWA', fontsize=16, fontweight='bold')
    
    # 1. Diagram Batang - Nilai Matematika
    axes[0, 0].bar(data['Nama'], data['Matematika'], color='orange', alpha=0.7)
    axes[0, 0].set_title('Nilai Matematika Siswa', fontweight='bold')
    axes[0, 0].set_xlabel('Nama Siswa')
    axes[0, 0].set_ylabel('Nilai')
    axes[0, 0].tick_params(axis='x', rotation=45)
    axes[0, 0].grid(True, alpha=0.3)
    
    # 2. Diagram Batang - Rata-rata Nilai
    axes[0, 1].bar(data['Nama'], data['Rata_rata'], color='skyblue', alpha=0.7)
    axes[0, 1].set_title('Rata-rata Nilai Siswa', fontweight='bold')
    axes[0, 1].set_xlabel('Nama Siswa')
    axes[0, 1].set_ylabel('Rata-rata')
    axes[0, 1].tick_params(axis='x', rotation=45)
    axes[0, 1].grid(True, alpha=0.3)
    
    # 3. Line Chart - Perbandingan Semua Mata Pelajaran
    x_pos = np.arange(len(data['Nama']))
    axes[1, 0].plot(x_pos, data['Matematika'], marker='o', label='Matematika', linewidth=2)
    axes[1, 0].plot(x_pos, data['Bahasa Indonesia'], marker='s', label='Bahasa Indonesia', linewidth=2)
    axes[1, 0].plot(x_pos, data['Informatika'], marker='^', label='Informatika', linewidth=2)
    axes[1, 0].set_title('Perbandingan Nilai Semua Mata Pelajaran', fontweight='bold')
    axes[1, 0].set_xlabel('Siswa')
    axes[1, 0].set_ylabel('Nilai')
    axes[1, 0].set_xticks(x_pos)
    axes[1, 0].set_xticklabels(data['Nama'], rotation=45)
    axes[1, 0].legend()
    axes[1, 0].grid(True, alpha=0.3)
    
    # 4. Pie Chart - Distribusi Rata-rata Nilai
    # Kategorikan nilai
    kategori = []
    for rata in data['Rata_rata']:
        if rata >= 90:
            kategori.append('Sangat Baik (≥90)')
        elif rata >= 80:
            kategori.append('Baik (80-89)')
        elif rata >= 70:
            kategori.append('Cukup (70-79)')
        else:
            kategori.append('Kurang (<70)')
    
    kategori_count = pd.Series(kategori).value_counts()
    colors = ['#ff9999', '#66b3ff', '#99ff99', '#ffcc99']
    axes[1, 1].pie(kategori_count.values, labels=kategori_count.index, autopct='%1.1f%%', 
                   colors=colors, startangle=90)
    axes[1, 1].set_title('Distribusi Kategori Nilai', fontweight='bold')
    
    plt.tight_layout()
    plt.show()

def buatGrafikTambahan(data):
    if data is None:
        return
        
    # Grafik Heatmap untuk melihat korelasi antar mata pelajaran
    plt.figure(figsize=(10, 6))
    
    # Subplot 1: Heatmap
    plt.subplot(1, 2, 1)
    correlation_matrix = data[['Matematika', 'Bahasa Indonesia', 'Informatika', 'Rata_rata']].corr()
    sns.heatmap(correlation_matrix, annot=True, cmap='coolwarm', center=0, 
                square=True, linewidths=0.5)
    plt.title('Korelasi Antar Mata Pelajaran', fontweight='bold')
    
    # Subplot 2: Box Plot
    plt.subplot(1, 2, 2)
    data_melted = pd.melt(data, id_vars=['Nama'], 
                         value_vars=['Matematika', 'Bahasa Indonesia', 'Informatika'],
                         var_name='Mata_Pelajaran', value_name='Nilai')
    sns.boxplot(data=data_melted, x='Mata_Pelajaran', y='Nilai')
    plt.title('Distribusi Nilai per Mata Pelajaran', fontweight='bold')
    plt.xticks(rotation=45)
    
    plt.tight_layout()
    plt.show()

def tampilkanStatistik(data):
    if data is None:
        return
        
    print("\n=== STATISTIK DESKRIPTIF ===")
    print(data[['Matematika', 'Bahasa Indonesia', 'Informatika', 'Rata_rata']].describe())
    
    print("\n=== RANKING SISWA ===")
    ranking = data.sort_values('Rata_rata', ascending=False).reset_index(drop=True)
    ranking.index += 1  # Mulai dari 1
    print(ranking[['Nama', 'Rata_rata']])

# Program Utama
if __name__ == "__main__":
    print("🎓 PROGRAM ANALISIS NILAI SISWA 🎓")
    print("="*50)
    
    # Jalankan analisis
    data_siswa = nilaiSiswa()
    
    if data_siswa is not None:
        # Tampilkan statistik
        tampilkanStatistik(data_siswa)
        
        # Buat visualisasi
        print("\n📊 Membuat visualisasi...")
        buatVisualisasi(data_siswa)
        
        print("\n📈 Membuat grafik tambahan...")
        buatGrafikTambahan(data_siswa)
        
        print("\n✅ Program selesai dijalankan!")
    else:
        print("\n❌ Program tidak dapat dijalankan karena ada masalah dengan data.")