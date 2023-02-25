import openpyxl
import pandas as pd
import datacompy
import xlsxwriter
from openpyxl import load_workbook
from datetime import date
from openpyxl.utils.dataframe import dataframe_to_rows

path = r"C:\Users\yaren\OneDrive\Masaüstü\de\MockStok.xlsx"
old_path = r"C:\Users\yaren\OneDrive\Masaüstü\de\MayısStokAsıl.xlsx"
xl = pd.ExcelFile('MayısStokAsıl.xlsx')
print("Sheet isimleri:", xl.sheet_names)

user = (input("Referans olarak almak istediğiniz ayın ismini giriniz"
              "(LÜTFEN BÜYÜK HARFLE YAZMAYINIZ!!!!!!):")).casefold()


for i in range(0, len(xl.sheet_names)):  #büyük harfle yazılınca ASCII den dolayı küçük harfe farklı geçiriyor
    if xl.sheet_names[i].casefold() == user.casefold():  #iki ismi de küçük harfli olarak kıyaslıyor
        user = xl.sheet_names[i]   #eğer eşitse sheet name in adını user a atıyor
        break

stock = pd.read_excel(path,sheet_name=0) #komple excel dosyasının ilk sheetini okumak için
stock_old = pd.read_excel(old_path,sheet_name = user)  #user artık sheet name le aynı karakterde

#stock.columns = [c.replace(' ', '_') for c in df.columns] #Boşluk yerine çizgi koy diyor
#stock.columns = stock.replace('Part_No','Part No')


stock.drop(stock.columns[[0,1,2]],axis=1,inplace=True) # birden fazla sütunu index üzerinden silmek
stock.drop(stock.columns[[3,4,5]],axis=1,inplace=True)
stock.drop(stock.columns[[4,5]],axis=1,inplace=True)

#stock ve old_stock ANA DATAFRAMELER
stock["Unnamed: 2"] = ""#boş sütun ekliyoruz
stock["Unnamed: 3"] = ""
stock["Unnamed: 4"] = ""

"""stock.insert(loc=2, column='Unnamed: 2', value=None)
stock.insert(loc=3, column='Unnamed: 3', value=None)
stock.insert(loc=4, column='Unnamed: 4', value=None)"""


stock = stock[["Part_No","Adı","Unnamed: 2","Unnamed: 3","Unnamed: 4","Stok","Bırımı","Stok*son Br Malıyet"]] #sıralamasını düzenlemek için
#stock.columns.values[0:2]=["kedi","patisi"] # 0. columndan 1. columna kadar isim değiştirmek için INDEX ÜZERİNDEN ÖRNEK
stock.rename(columns={stock.columns[7]:"Toplam_Maliyet"},inplace=True) #index üzerinden tekil satır değiştirmek için, yeniden isimlendirme

isolated_oldstock = pd.DataFrame()
isolated_stock = pd.DataFrame()
#kedi = stock.loc[stock['Adı'] =="EMAS PEDAL PDKS11BX10" ] # ÖRNEK, Sadece "EMAS PEDAL" değerini sağlayanlardan dataframe oluşturuyor
isolated_oldstock=stock_old.filter(["Part_No","Stok","Toplam_Maliyet"]) #izole etmek için, sadece part no ve stoklardan yeni bir dataframe oluşturuyoruz
isolated_stock=stock.filter(["Part_No","Stok","Toplam_Maliyet"])
#print(isolated_oldstock)
#print(isolated_stock)

#sanırım yeniden eskiye yaparken how'ı right seçersek yeni eklenenleri de bulabiliriz
oldtonew=isolated_oldstock.merge(isolated_stock.drop_duplicates(),on=["Part_No"],how="left",indicator=True) # !!!Left_only yazanlar oldda olup newde olmayanlar, çünkü old sonda kaldı
#print("OLDTONEW",oldtonew.to_string()) #yukarıdaki indicator bu tip yapılarda bayağı yardımcı
newtoold=isolated_stock.merge(isolated_oldstock.drop_duplicates(),on=["Part_No"],how="left",indicator=True) #left only yazanlar newde olup oldda olmayanlar,çünkü new solda kaldı

oldtonew_memb = oldtonew[oldtonew["_merge"]=="left_only"] #ÖNEMLİ ÖRNEK, MERGE DEĞERİ BULUNAN COLUMNDA LEFT ONLY OLANLARI FİLTRELİYORUZ VE YENİ DATAFRAME OLUŞTURUYORUZ
#left_only demek eskide var ama yeni dataframede yok demek, diğer dataframe için de aynısı geçerli,İZOLE EDİP AYIRIYORUZ BÖYLECE HESAPLAYABİLECEĞİZ
newtoold_memb = newtoold[newtoold["_merge"] == "left_only"]
print("SDKASJFSKDFHNDSJBHBSDGFBJSDFSDHF****************",newtoold_memb)


olddf_list = oldtonew_memb["Part_No"].tolist() # Part no kısmının listeye çevrilmiş hali, eski ayın stoğu, eski ayda olup yenide olmayanlar, stoktan düşen elemanlar
print(oldtonew_memb)
print("Eski ayda olup yeni ayda olmayanlar,stoktan düşen",olddf_list)

newdf_list = newtoold_memb["Part_No"].tolist() # yeni ayın stoğunun listeye dönüşümü, yeni ayda olup eski ayda olmayan, stoğa eklenen elemanlar
print(newtoold_memb)
print("Stoğa eklenen",newdf_list)
#NOT PART NONUN YANINA STOK DA EKLERSEK STOK DEĞİŞİMLERİNİ GÖRÜRÜZ, STOK DEĞİŞİMİ İÇİN OLAN KISIM, STOK VEYA MALİYETİ DEĞİŞENLER, beraber veya ayrı ayrı değişse de olur
#DATAFRAME OLDULAR
stk_newtoold=isolated_stock.merge(isolated_oldstock.drop_duplicates(),on=["Part_No","Stok","Toplam_Maliyet"],how="left",indicator=True) # stok ve toplam maliyet de yazınca değişim burada da var mı anlıyoruz
stk_memb = stk_newtoold[stk_newtoold["_merge"]=="left_only"] #sadece değişim oluşmuş olanları GÖSTERMEK İÇİN
stk_memb_list = stk_memb["Part_No"].tolist() #değişenlerin part nolarıyla listesi, burada tek sıkıntı yeni eklenenleri de ekliyor
print("Stok veya Maliyetleri Değişenler + yeni eklenenler",stk_memb.to_string()) # yeni eklenenlerin stok ve maliyetini yenilememizin etkisi olmuyor
print("LİSTE Stok veya Maliyetleri Değişenler + yeni eklenenler LİSTE",stk_memb_list)

'''oldtonew=isolated_oldstock.merge(isolated_stock.drop_duplicates(),on=["Part_No","Stok"],how="right",indicator=True)
print("Sanırım stoktan düşen",oldtonew.to_string()) # buradaki bir üstüyle aynı, yenide olup eskide olmayanları ayırıyor ama bu sefer how kısmına right dedik'''



#oldtonew_memb --- eski stokta olup yeni stokta olmayan, newtoold ise yenide olup eskide olmayan(yani eklenen), İZOLE EDİP AYIRIYORUZ BÖYLECE HESAPLAYABİLECEĞİZ
#!!!!!NOT STOK SAYILARINI DA KONTROL ETMEK GEREKEBİLİR Mİ!!!!!
#NOT 2, NORMAL ŞARTLARDA ANA DATAFRAMEİ KÜÇÜLTMEDEN DE BU VERİLERİ KARŞILAŞTIRABİLİRİM SANIRIM


stok_eklenen=pd.DataFrame()
stok_dusen=pd.DataFrame()
#yeni eklenen ve stoktan çıkanları dataframelerde bulup iki ayrı dataframe'e çevireceğiz, sonradan eklemek için ayrı dosyaya
for i in range(0,len(newdf_list)):
    kedi = stock.loc[stock['Part_No'] ==newdf_list[i]] # büyük dataframe içerisinde bir üyeyi arayıp o üyedenin elemanlarından dataframe yapmak için
    #stok_eklenen=kedi.append(stok_eklenen,ignore_index=True)
    stok_eklenen=pd.concat([stok_eklenen,kedi],axis=0) #!!!


for j in range(0,len(olddf_list)):
    pati = stock_old.loc[stock_old["Part_No"] == olddf_list[j]] #Part No kısmında old_dflist deki elemanları dahil olanları dataframe yapıyoruz(append ederek tabii)
    #stok_dusen=pati.append(stok_dusen,ignore_index=True)
    stok_dusen=pd.concat([stok_dusen,pati],axis=0) #!!!
    # Part numarası üzerinde newdf_list(yani yeni dataframede bulunup eskide bulunmayanlar) elemanlarını aldık
    #BU YENİ LİSTEYİ EXCELE TEKRARDAN MI YAZMALI YOKSA SADECE DEĞERLERİNİ Mİ TOPLAMALI
    #sütun ve satır maksimum sayısını alacağım, ana stok dosyasının maksimum satırının altında yeni eklenen elemanları highlight mı etmeli
    #kedinin farklı versiyonları append edip yeni dataframe oluşturabilirim
print("EKLENEN STOK",stok_eklenen)
print("DÜŞEN STOK",stok_dusen)


#new_mainstock=newtoold #neden koydum hatırlamıyorum ama sanırım bir işe yaramıyor
#BÜYÜK DOSYA OLUŞTURMAK İÇİN#
new_mainstock=stock_old.merge(stock.drop_duplicates(),how="left",indicator=True)#eski ana stokla yeniyi birleştiriyoruz
#left ile sadece eski stokta olanları ve ortak olanları dahil etmiş oluyoruz, fakat stoktan düşenleri manuel düşürmemiz gerekecek, eklenenler de manuel eklenecek

#stoktan düşenleri silmek için olddf_list ve oldtonew_memb kullanılabilir, stoktan düşenler
for x in range(0,len(olddf_list)):
    new_mainstock=new_mainstock[new_mainstock.Part_No !=olddf_list[x]] # SİLMEK İÇİN!, elimizde stoktan düşenlerin listesi zaten vardı olddf_list, buradan her bir elemanını gösterip siliyoruz
#new_mainstock=stok_eklenen.append(new_mainstock,ignore_index=True)
new_mainstock=pd.concat([new_mainstock,stok_eklenen],axis=0)
# eklenenlerin dataframeini ekliyoruz #BURA SIKINTI ÇIKARABİLİR!!!!
#new_mainstock= pd.concat([new_mainstock,stok_eklenen],axis=0) #concat kullandığımda problem yaşıyorum şu an!!!

#print("YENİ Mİ",new_mainstock.to_string())

indexi = []
rownum = 0
for y in range(0,len(stk_memb_list)):
    #ÇALIŞAN BU, merge yapısı sebebiyle stok ve maliyet değişenlerde eski dosyayı referans alıyordu, buradan onu değiştiriyoruz
    to_be_replaced_maliyet = (stk_memb.loc[stk_memb["Part_No"] == stk_memb_list[y]]).iloc[0]["Toplam_Maliyet"] # maliyeti değişenlerin yeni maliyet değerini spesifik olarak almak
    degerler = new_mainstock.loc[new_mainstock["Part_No"] == stk_memb_list[y]] #buradaki degerler ayrı bir dataframe olduğu için ana dosyaya ETKİSİZZZ!!! BURADAN DEĞİŞİKLİK YOK
    #degerler üzerinden değişiklik yok fakat DEĞİŞTİRMEMİZ GEREKEN İNDEX LİSTESİNİ BURADAN ELDE EDEBİLİRİZ. YANİ DEĞİŞTİRMEMİZ GEREKEN İNDEXLERİ BURADAN BULDUK
    index_eleman= degerler.index.tolist() #kullanacağımız elemanların indexlerini ayırıyoruz teker teker
    indexi.append(index_eleman[0]) # asıl listede append ediyoruz indexleri
    new_mainstock.loc[indexi[rownum],"Toplam_Maliyet"] = to_be_replaced_maliyet  # farklı indexleri olduğu için index listesine ihtiyacımız var, eleman değiştirmek için
    #ELEMAN DEĞERİ DEĞİŞTİRME LOC İLE FAKAT İNDEX GEREKİYOR, STOK DEĞİŞİMİ DE YAPILMALI
    rownum=rownum+1
    #print(degerler.to_string())

    '''main_oldvalue_stock = (new_mainstock.loc[new_mainstock["Part_No"]==stk_memb_list[y]]).iloc[0]["Stok"] #direkt olarak değerleri almak için ama artık kullanmıyorum
    main_oldvalue_maliyet = (new_mainstock.loc[new_mainstock["Part_No"] == stk_memb_list[y]]).iloc[0]["Toplam_Maliyet"]'''


uniqueValues=new_mainstock["Unnamed: 2"].drop_duplicates().unique() # unik değerleri kendilerini tekrar etmeyecek şekilde almak
print("UNİİİİQQEEEEE",uniqueValues)

rownum=0
ms_colnum = new_mainstock.shape[1]  # 1 sütun sayısı, 0 satır sayısını gösteriyor , ms = mainstock col number
col_list = new_mainstock.index.tolist() #dağıtınık index olduğu için index listesini alıyoruz
ms_colnum+=1# son sütunu aşsın diye
bos=0
for o in range(0,len(uniqueValues)):
    if uniqueValues[o] == "":  # boş elemanları ayırmak için
        boslar = new_mainstock.loc[new_mainstock["Unnamed: 2"] == uniqueValues[o]]  # Unnamed: 2 olan sütnunun uniqueValues[t] dizisinin elemanı olanını buluyoruz
        # boslar dataframei ayrı dataframe olduğu için asıl dataframe e etkisiz ama buradan index listesini çıkarabiliriz
        boslar_index = boslar.index.tolist()  # boş elemanlı dataframein index listesini çıkarıyoruz(Unnamed: 2),

        for bos in range(bos, len(boslar_index)):  # boş olanlara inputla isim ekleyeceğiz # BOS YERİNE 0 YAPIP DENE

            degisecek_isim = boslar.loc[boslar_index[bos], "Adı"],
            print(degisecek_isim)
            inputumuz = input("isimli parçanın nereye ait olduğunu giriniz:")
            new_mainstock.loc[boslar_index[bos], "Unnamed: 2"] = inputumuz
            # print("HEYEYEYEYE",new_mainstock.loc[boslar_index[bos],"Unnamed: 2"])

uniqueValues = new_mainstock["Unnamed: 2"].unique()

for t in range(0,len(uniqueValues)):
    #print("HEHYEYEY",boslar.to_string())
    #uniqueValues = new_mainstock["Unnamed: 2"].unique() #input girdikten sonra unik elemanlar değişebileceği için bir daha alıyoruz
    #print("YENİUNİQEEEEE",uniqueValues)
    unik_elemanlar = new_mainstock[new_mainstock["Unnamed: 2"] == uniqueValues[t]] # unik elemanlardan oluşan bir dataframe oluşturuyor, str contains kullanınca sanırım yanlış
    toplam = unik_elemanlar["Toplam_Maliyet"].sum() # o elemanların toplam maliyetten toplamını alıyor,
    #print(unik_elemanlar.to_string())
    print("TOPLAMI",uniqueValues[t],toplam)
    new_mainstock.loc[col_list[rownum], ms_colnum] = toplam
    new_mainstock.loc[col_list[rownum], ms_colnum + 1] = str(uniqueValues[t])
    rownum = rownum +1
print("YENİUNİQEEEEE",uniqueValues)
  # 1 sütun sayısı, 0 satır sayısını gösteriyor , ms = mainstock col number, BİR DAHA ALIYORUZ BU SEFER DİĞER DATAFRAMELERİ EKLEYECEĞİZ
#new_mainstock=new_mainstock.sort_values(by,axis=0,ascending=False) # sort ediyoruz

'''print("Eski ayda olup yeni ayda olmayanlar,stoktan düşen",olddf_list)
print("Stoğa eklenen",newdf_list) LİSTE#liste olarak '''
# stoktan düşeni direkt olarak alabiliriz çünkü nereye ait olduğu(plastik atölye vs) belirli, stoğa eklenenleri ise en son nereye ait olduklarını input etmiştik
# o sebeple tekrardan almamız gerekecek liste yardımıyla
#stk_memb_list stoktan düşen veya maliyeti değişenler liste halinde, YENİ EKLENENLER DAHİL OLARAK
new_mainstock.drop("_merge", inplace=True, axis=1)


df_ek=pd.DataFrame()
df_ek_selection=pd.DataFrame()
abc=pd.DataFrame()
abc_ek=pd.DataFrame()
for l in range(0,len(stk_memb_list)):# "stok değeri değişen veya maliyeti değişen" komple listeyi almak için, bu noktada yeni eklenenler de dahil
    abc = new_mainstock.loc[new_mainstock["Part_No"] == stk_memb_list[l]]
    #abc_ek = abc.append(abc_ek,ignore_index=True)# abc_ek stk_membin aynısı ama adı kısmı ve unnamed kısımları var, stk membde orası yoktu
    abc_ek = pd.concat([abc_ek, abc], axis=0)
    # stok_dusen = pd.concat([stok_dusen, pati], axis=0)  # !!! stok düşene pati ekleniyor ÖRNEK CONCAT
for b in range(0,len(newdf_list)): # stoğa eklenenlerin isimlerini almak için, inputla manuel girmiştik
    eklenen = new_mainstock.loc[new_mainstock["Part_No"] == newdf_list[b]] # bir daha alıyoruz çünkü sonradan nereye ait olduklarını elle girdik, ilk başta belirsizdi
    #df_ek = eklenen.append(df_ek,ignore_index = True)
    df_ek=pd.concat([df_ek,eklenen],axis=0)
    df_ek_selection = df_ek[["Part_No","Adı","Unnamed: 2","Stok","Bırımı","Toplam_Maliyet"]] # istediğimiz kısımları almak için # filtre de kullanılabilir
print("Stokğa EKLENEN",df_ek_selection.to_string())
print("Stok düşen", stok_dusen.to_string())

#print("EKLENEN VE MALİYET/STOK DEĞİŞEN LİSTESİ",stk_memb.to_string())
stk_memb=abc_ek # abc_ek stk_membin aynısı ama adı kısmı ve unnamed kısımları var
deger_degisen=stk_memb
#print("SDŞLFKDKSJFHKSDKFHDS", newdf_list)
for n in range(0,len(newdf_list)):
    deger_degisen=deger_degisen[deger_degisen.Part_No !=newdf_list[n]] #stoğa eklenenler ve değeri değişenler aynı yerde kalıyordu, eklenenleri siliyoruz
    # nedense deger_degisen=stk_memb[stk_memb.Part_No !=newdf_list[n] yapınca olmuyor
# eklenen stok ve değişen stokların bütün olduğu yerden ekleneni çıkarıp sadece değişenleri de alabiliriz.
#print("!!!!Stok değeri veya maliyeti değişenler",deger_degisen.to_string())
print("STOK DEĞERİ VEYA MALİYET DEĞİŞEN LİSTESİ YENİ EKLENEN YOK!",deger_degisen.to_string())
print("STOĞA EKLENEN(yeni eklenen) LİSTE",newdf_list)

df_ek_selection.reset_index(drop=True,inplace=True)# indexleri silmezsek concat yapamıyor NaN değeri veriyor
deger_degisen.reset_index(drop=True,inplace=True)
stok_dusen.reset_index(drop=True,inplace=True)

#del deger_degisen["Unnamed: 3"]
#del deger_degisen["Unnamed: 4"] # silmek için alternatif

deger_degisen.drop("Unnamed: 3",inplace=True,axis=1) #silmek için
deger_degisen.drop("Unnamed: 4",inplace=True,axis=1) #İNPLACE YAPMAZSAK ESKİ HALİNİ RETURN EDİYOR

#df_ek_selection.columns=["Part_No_Stk_Eklenen","Adı","Unnamed: 2","Stok","Bırımı","Toplam_Maliyet"] #isim değişikliği için KOMENTE ALDIK, kommentte olmayınca HATA ALINIYOR
deger_degisen.rename(columns={"Part_No":"Part_No_Stk_Deger_Maliyet_Degisen"},inplace=True)#alternatif isim değişikliği
stok_dusen.rename(columns={"Part_No":"Part_No_Stk_Dusen"},inplace=True)
#deger_degisen.columns=["Part_No_Stk_Deger_Maliyet_Degisen","Adı","Unnamed: 2","Stok","Bırımı","Toplam_Maliyet"]
#stok_dusen.columns=["Part_No_Stk_Dusen","Adı","Unnamed: 2","Stok","Bırımı","Toplam_Maliyet"]
degisenler =pd.concat([df_ek_selection,deger_degisen,stok_dusen],axis=1)


new_mainstock.rename(columns={"10":"toplam_b"},inplace=True)  #header yazdıramadım
new_mainstock.rename(columns={"11":"bölümler"},inplace=True)  #header yazdıramadım
print(degisenler.to_string())
'''
#dataframe olarak
print("EKLENEN STOK",stok_eklenen.to_string())
print("DÜŞEN STOK",stok_dusen.to_string())'''

#second_path = r"C:\Users\kayaa\PycharmProjects\Stok\Lazer.xlsx"



'''common = isolated_oldstock.merge(isolated_stock,on=['Part_No','Stok'],how="right",indicator=True)
print("HEYEYEYEY",common.to_string()) #inner otuer farklarını anlamak için örnek, BURADAN ANA DOSYAYI MERGELEYEBİLİRİM'''



FilePath = r"C:\Users\yaren\OneDrive\Masaüstü\de\MayısStokAsıl.xlsx"
ExcelWorkbook = load_workbook(FilePath)
writer1 = pd.ExcelWriter(FilePath, engine='openpyxl')
writer1.book = ExcelWorkbook

dt = str(date.today().month)  #sheet adını yazdırmak için, referans olarak bu ayı alıyor

if dt == '2':
    dt = 'Ocak'
elif dt == '3':
    dt = 'Şubat'
elif dt == '4':
    dt = 'Mart'
elif dt == '5':
    dt = 'Nisan'
elif dt == '6':
    dt = 'Mayıs'
elif dt == '7':
    dt = 'Haziran'
elif dt == '8':
    dt = 'Temmuz'
elif dt == '9':
    dt = 'Ağustos'
elif dt == '10':
    dt = 'Eylül'
elif dt == '11':
    dt = 'Ekim'
elif dt == '12':
    dt = 'Kasım'
else:
    dt = 'Aralık'

new_mainstock.to_excel(writer1, sheet_name=dt)  #dataframei excele kaydetmek için
degisenler.to_excel(writer1, sheet_name= dt + 'değişenler')  #dataframei excele kaydetmek için


writer1.save()

