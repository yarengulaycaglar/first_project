import pandas as pd
import numpy as np
from datetime import datetime
import matplotlib.pyplot as plt
from matplotlib import interactive
import seaborn as sns
import random
import PySimpleGUI as sg


#basit bir GUI oluşturdum
sg.theme('SandyBeach')
layout = [
    [sg.Text('Arıza dosyasının konumunu giriniz:')],
    [sg.Text('Dosya', size =(15, 1)), sg.InputText()],
    [sg.Submit(), sg.Cancel()]
]
window = sg.Window('Arıza', layout)
event, values = window.read()
window.close()



colors = ["#d7e1ee", "#cbd6e4", "#bfcbdb", "#b3bfd1", "#a4a2a8", "#df8879", "#c86558", "#b04238", "#991f17",
          "#3b3734", "#474440", "#54504c", "#6b506b", "#ab3da9", "#de25da", "#eb44e8", "#ff80ff",
          "#ffb400", "#d2980d", "#a57c1b", "#786028", "#363445", "#48446e", "#5e569b", "#776bcd", "#9080ff",
          "#e27c7c", "#a86464", "#6d4b4b", "#503f3f", "#333333", "#3c4e4b", "#466964", "#599e94", "#6cd4c5"]  #renk kodları değiştirilebilir



df = pd.read_excel(io = values[0])
df.columns = [c.replace(' ', '_') for c in df.columns]  #sütun adlarında boşluk varsa onları _ la değiştir

df = df.filter(['BE_No', 'Nesne_No', 'Nesne_Açıklaması', 'Çalışma_Detayı','Hata_Tanımı', 'Ortalama_Çalışma_Saati',
                'Fiili_Başlangıç','Fiili_Bitiş', 'Öncelik', 'Hata_Türü_Tanımı'])

df.insert(loc=2, column='Bölgeler', value=None)  #Bölgeler adında boş bir sütun açtım

df_database = pd.read_excel(r"C:\Users\yaren\OneDrive\Masaüstü\DataBaseArza.xlsx")
df_database.columns = [c.replace(' ', '_') for c in df_database.columns]  #headerlarda isim arasında boşluk olanları '_' la değiştirdim
nesne_no = df['Nesne_No'].values
nesne_no_data_base = df_database['Nesne_No'].values
bolgeler_data_base = df_database['Bölgeler'].values

nesne_no_list = nesne_no.tolist()
nesne_no_data_base_list = nesne_no_data_base.tolist()
bolgeler_data_base_list = bolgeler_data_base.tolist()
bolge=[]

"""for i in range(0,len(bolgeler_data_base_list)):  #ŞUANDA İŞİMİZE YARAMIYOR AMA İLERİDE İNDEXİ ALMAK İÇİN KULLANABİLİRİZ
    degerler = df.loc[df["Nesne_No"]==nesne_no_data_base_list[i]]
    bolge.append(bolgeler_data_base_list[i])
    #print("DEGERLER",degerler)
    degerler_index=degerler.index.to_list()
    for x in range(0,len(degerler_index)):
        indexi.append(degerler_index[x])
rownum=0
print("****",bolge)
for z in range(0,len(indexi)):
    df.loc[indexi[rownum],"Nesne_No"]=bolge[z]
    rownum=rownum+1
    #print("DEGERLERINDEXXXXX",degerler_index)
#print("YOVVV",indexi)
#print("YOYOYOYOOY",degerler)"""

for i in range(0, len(nesne_no_data_base_list)):
    df.loc[df['Nesne_No'] == nesne_no_data_base_list[i], 'Bölgeler'] = bolgeler_data_base_list[i]
    #dataframe deki Nesne no eğer data basedeki nesne noya eşitse(data base in nesne nosunu liste olarak aldım), dataframedeki bölgeler sütununa
    #data base deki bölgeler sütunundaki datayı yaz(data base deki bölgeler sütununu liste olarak aldım)

new_bolgeler_index = df['Bölgeler'].index.tolist()
new_bolgeler_list = df['Bölgeler'].values.tolist()

add_NesneNo_Bolgeler_temp = [] #geçici liste
add_NesneNo_Bolgeler = []
# elemanı olmayanları girmek için
for i in range(0, len(new_bolgeler_list)):
    if new_bolgeler_list[i] == None:  #eğer bölgelerde değer yoksa
        print("Nesne No:", nesne_no_list[i])
        user = input("Bölgeler:")
        df.loc[new_bolgeler_index[i], 'Bölgeler'] = user  #kullanıcıdan nesne numarasının ait olduğu bölgeyi al
        #data basedeki verileri de kullanarak ana dosyamızdaki bölgeler sütununu doldurduktan sonra kalan boş yerleri manuel olarak doldurmak için
        #bu döngüyü kullanıyoruz
        #her bir None ın indexini kullanarak kullanıcıdan aldığımız veriyi o bölgeler sütununa idexdeki satıra giriyoruz
        add_NesneNo_Bolgeler_temp.append(nesne_no_list[i])  #nesne numarasını geçici listeye ekle
        add_NesneNo_Bolgeler_temp.append(user)  #bölgesini geçici listeye ekle
        add_NesneNo_Bolgeler.append(add_NesneNo_Bolgeler_temp)  #nesne no ve bölgenin olduğu listeyi ana listeye ekle
        add_NesneNo_Bolgeler_temp = []  #tekrardan her bir eleman yazılmasın diye listeyi sıfırla
df_database_add = pd.DataFrame(add_NesneNo_Bolgeler, columns=['Nesne_No', 'Bölgeler'])  #listeyi nesne no ve bölgeler headerları altın dataframe olarak al
df_database = pd.concat([df_database,df_database_add], axis=0)  #data base imize yeni gelen nesne no larının olduğu datta framei ekle
df_database = df_database.sort_values(by=['Bölgeler'])  #Bölgelere göre sırala
df_database.drop('Unnamed:_0', inplace=True, axis=1)  #indexi sıfırladıkttan sonra eski index, index kolonu altında kalıyor onu silmek içib kullanıyoruz
df_database.to_excel('DataBaseArza.xlsx')  #data basemiz güncellendi"""

df_1 = df.loc[df['Öncelik'] == 1]  #ASIL İHTİYACIMIZ OLAN DURUŞLU ARIZALAR

#bölgeleri ayır her bir bölgenin ortak çalışma saatini topla bunu 'Total' headerının altına yazdır
df_1_total = df_1.groupby(['Bölgeler'])['Ortalama_Çalışma_Saati'].sum().reset_index(name='Total')  #excel dosyasına virgüllü şekilde yazdırıyor!!!
df = pd.concat([df,df_1_total], axis=1)  #ana dataframe mimize ortalama çalışma saatinin toplamını yazdırdığımız data framei ekle






#Ortalama çalışma saati en yüksek olan il 5 tanesini(1lerden) al bir yere Bölgeler, Nesne açıklaması, Çalışma detayını yazdır
df_max = df_1.iloc[df_1['Ortalama_Çalışma_Saati'].argsort()[-5:]].reset_index()
df_max = df_max.filter(['Bölgeler','Nesne_Açıklaması', 'Çalışma_Detayı', 'Ortalama_Çalışma_Saati'])
df = pd.concat([df,df_max], axis=1)  #ana dataframe mimizle ortalama çalışma saatinin toplamını yazdırdığımız data framei topla

df_toplam_arza_sayisi = df_1.groupby(['Bölgeler']).size().reset_index(name='Topam arıza sayısı')  #excel dosyasına virgüllü şekilde yazdırıyor!!!
df = pd.concat([df,df_toplam_arza_sayisi], axis=1)  #ana dataframe mimizle ortalama çalışma saatinin toplamını yazdırdığımız data framei topla
df.insert(loc=11, column='11', value=None)
df.insert(loc=14, column='14', value=None)
df.insert(loc=18, column='19', value=None)




#burdan sonra arıza saatlarinin yoğunluğunu bulmaya çalışacağız!!!
df_yogunluk_saat = df.groupby(pd.to_datetime(df["Fiili_Başlangıç"]).dt.hour).size().reset_index(name='Yoğunluk_Saat')
#df üzerinden fiili başlangıcın saaatini al, yoğunluğunu bul, yoğunluk_saat olarak kaydet
df_yogunluk_gun = df.groupby(pd.to_datetime(df["Fiili_Başlangıç"]).dt.day).size().reset_index(name='Yoğunluk_Gün')
#df üzerinden fiili başlangıcın gününü al, yoğunluğunu bul, yoğunluk_saat olarak kaydet

#Burası hangi günlerde daha çok arıza çıktığını görselleştirmek için
df_yogunluk_gun.plot.bar(x="Fiili_Başlangıç", y="Yoğunluk_Gün")
sns.barplot(data=df_yogunluk_gun, x="Fiili_Başlangıç", y="Yoğunluk_Gün")
sns.set(rc={'figure.figsize':(15,10)})
plt.xlabel('Gün')
plt.ylabel('Arıza Sayısı')
plt.title('Arıza Sayısının Günlere Göre Yoğunluğu', size=18)
plt.xlabel('')
plt.legend(title="Gün" ,loc=1)
plt.xticks(size=16, rotation=90)
plt.yticks(size=16, rotation=90)
plt.rcParams['figure.dpi'] = 360
plt.show()


#Burası hangi saatlerde daha çok arıza çıktığını görselleştirmek için
df_yogunluk_saat.plot.bar(x="Fiili_Başlangıç", y="Yoğunluk_Saat")
sns.barplot(data=df_yogunluk_saat, x="Fiili_Başlangıç", y="Yoğunluk_Saat")
sns.set(rc={'figure.figsize':(15,10)})
plt.xlabel('Saat')
plt.ylabel('Arıza Sayısı')
plt.title('Arıza Sayısının Saatlere Göre Yoğunluğu', size=18)
plt.xlabel('')
plt.legend(title="Saat" ,loc=1)
plt.xticks(size=16, rotation=90)
plt.yticks(size=16, rotation=90)
plt.rcParams['figure.dpi'] = 360
plt.show()



#en yükske arıza olan üç saat alınacak, bu saatlerde olan arızalar bölgeye bölünecek
#bölgelerin arıza miktarının nasıl dağıldığını pie üzerinde yüzdeyle göstersin
df_son = pd.DataFrame()
df_max_saat_temp = df_yogunluk_saat.iloc[df_yogunluk_saat['Yoğunluk_Saat'].argsort()[-3:]].reset_index(drop=True)  #gün içinde en çok arıza alan üç saat
max_saat_list = df_max_saat_temp["Fiili_Başlangıç"].tolist()




#Burası veriyi görselleştirmek için  BU GRAFİK EN ÇOK ARIZA ÇIKAN 3 SAAT İÇİN!!!!!!!
df_son = pd.DataFrame()
df_temp = pd.DataFrame()
for a in range(0,len(max_saat_list)):
    df_temp =df.loc[df["Fiili_Başlangıç"].dt.hour == max_saat_list[a]]
    df_son = pd.concat([df_son,df_temp],axis=0)
df_son = df_son.T.drop_duplicates().T  #tekrar edenleri silsek bile 2 tane bölgeler sütunu oluşuyor
df_son = df_son.iloc[:,0:11 ]  #Bunun önüne geçmek için tekrar edenleri siliyoruz kalanlardan ilk 10 sütunu alıyoruz

"""#Bölgelere ayır, hangi saatte kaç kere arızalandığını say, Bölgelerin_Yoğunluğu sütunu oluştur ve kaydet
df_yogunluk_saat_bolgeler = df_son.groupby(['Bölgeler', df_son['Fiili_Başlangıç'].dt.hour]).size().reset_index(name='Bölgelerin_Yoğunluğu')
bolgelerin_adlari = df_yogunluk_saat_bolgeler['Bölgeler'].T.drop_duplicates().T.tolist()  #bölgelerin adlarını liste olarak alıyor
x = []  #bu boş listeye bölgelerin adları kadar 0 ekleyeceğiz, bunu sonradan slice büyütmek için kullanacağız
for i in range(0, len(bolgelerin_adlari)):
    x.append(0)
for i in range(0, len(bolgelerin_adlari)):  #döngü yaptık çünkü her bir bölge adı için ayrı grafik oluşturacağız
    data = df_yogunluk_saat_bolgeler.loc[df_yogunluk_saat_bolgeler['Bölgeler'] == bolgelerin_adlari[i]]  #Burda her bir bölge adı için dataframe oluşturduk
    data_set = data['Bölgelerin_Yoğunluğu'].to_numpy()  #arıza sayısını numpy la array olrak aldık çünkü pie içinde kullanacağız
    my_labels = data['Fiili_Başlangıç'].tolist()  #saatleri list olarak aldık çünkü pie içinde kullanacağız
    def func(pct):  #yüzde göstermek için
        return "{:1.1f}%".format(pct)
    plt.pie(data_set, labels=my_labels,  autopct=lambda pct: func(pct), shadow=True)
    plt.legend(title='Arıza saati')
    plt.title('Arızaların Bölgelere göre Saatlik Yoğunluğu\nEn Çok Arıza Çıkan İlk Üç saat için: {}'.format(bolgelerin_adlari[i]))
    plt.axis('equal')
    plt.show()
"""

##Burası veriyi görselleştirmek için  BU GRAFİK TÜM SAATLER İÇİN!!!!!!!
df_temp3 = df.iloc[:,0:11 ]  #Aynı sütun adlarına sahip birçok sütun var,
# hata almanın önüne geçebilmek için ana dataframe in ilk 10 sütunun aldım(başlangıçta da 10 sütun vardı)
#ana dataframei bölgelere ayırıp bu bölgelerin arıza başlangıç saatlerini alarak saatlik yoğunluklarını bulduk,
# bunu ayrı bir dataframe üzerinden Bölgelerin_Yoğunluğu_Tüm
#sütunu altında kaydettik

df_yogunluk_saat_bolgeler_tum = df_temp3.groupby(['Bölgeler', (df['Fiili_Başlangıç'].dt.hour).astype(str)+(":00")]).size().reset_index(name='Bölgelerin_Yoğunluğu_Tüm')
#str olarak cast ettik çünkü :00 ı diğer türlü ekleyemiyorduk
bolgelerin_adlari_tum = df_yogunluk_saat_bolgeler_tum['Bölgeler'].T.drop_duplicates().T.tolist()  #bölgelerin adlarını liste olarak alıyor
#Bölgelere ayır, hangi saatte kaç kere arızalandığını say, Bölgelerin_Yoğunluğu_Tüm sütunu oluştur ve kaydet
for i in range(0, len(bolgelerin_adlari_tum)):  #döngü yaptık çünkü her bir bölge adı için ayrı grafik oluşturacağız
    data = df_yogunluk_saat_bolgeler_tum.loc[df_yogunluk_saat_bolgeler_tum['Bölgeler'] == bolgelerin_adlari_tum[i]]  #Burda her bir bölge adı için dataframe oluşturduk
    data_set = data['Bölgelerin_Yoğunluğu_Tüm'].to_numpy()   #arıza sayısını numpy la array olrak aldık çünkü pie içinde kullanacağız
    if len(data_set) > 2:  #arıza sayısı 2 den fazla olduğu zaman pie yapar
        sum = data_set.sum()  # toplam kaç arıza olden yazdırmak için
        my_labels = data['Fiili_Başlangıç'].tolist()  #saatleri list olarak aldık çünkü pie içinde kullanacağız

        def func(pct):  # yüzde göstermek için
            return "{:1.1f}%".format(pct)

        wp = {'linewidth': 2, 'edgecolor': "white"}
        plt.pie(data_set, labels=my_labels, autopct=lambda pct: func(pct), radius=1, labeldistance=1.03,
                colors=random.sample(colors, len(bolgelerin_adlari_tum)), textprops={'fontsize': 15, 'weight':'bold'},
                 frame=False,  wedgeprops=wp, rotatelabels=False, startangle=90)
        plt.legend(labels=my_labels, fontsize=10, loc='upper right')
        plt.legend(title='Arıza saati', fontsize='medium')
        plt.title("Arızaların Bölgelere göre Saatlik Yoğunluğu: {}\nToplam arıza sayısı: {}".format(bolgelerin_adlari_tum[i], sum),
                  color="black", fontsize=20, fontweight= "bold", x=0.5, y=1.06)
        plt.axis('equal')
        plt.savefig("{}.png".format(bolgelerin_adlari_tum[i]), dpi=300)
        plt.show()


"""#BURADAN SONRASI "BÖLÜMLER BAZINDA DURUŞ SÜRESİ(%)-AY" GÖRSELLEŞTİRMEK İÇİN
set_data = df_1_total['Total'].to_numpy()  # Total numpy la array olrak aldık çünkü pie içinde kullanacağız
labels = df_1_total['Bölgeler'].tolist()  # bölgeleri list olarak aldık çünkü pie içinde kullanacağız
explode = []
for i in range(0,len(labels)):
    explode.append(0)
x = 0.02
for i in range(0, len(labels)):
    if set_data[i] < 10:
        explode[i] = x
        x += 0.07
explode = tuple(explode)
print(explode)
wp = {'linewidth': 5, 'edgecolor': "white"}
plt.pie(set_data, colors=random.sample(colors, len(set_data)), labels=labels, radius=1,
        autopct='%1.1f%%', pctdistance=1.1,  labeldistance=1.2, textprops={'fontsize': 11, 'weight':'bold'},
        explode=explode, frame=False,  wedgeprops=wp, rotatelabels=False, startangle=0)

centre_circle = plt.Circle((0, 0), 0.70, fc='white')
fig = plt.gcf()
fig.gca().add_artist(centre_circle)
plt.title('Bölümler Bazında Duruş Süreleri(%) - AY:',  color="black", fontsize=22, fontweight= "bold", x=0.5, y=1.07)
plt.legend(loc='lower right',fontsize=8)
plt.legend(title='Duruş Süreleri', fontsize='medium')
plt.savefig("duruş_süreleri.png", dpi=300)
plt.show()
"""


df.to_excel('Arıza_Yeni.xlsx')




