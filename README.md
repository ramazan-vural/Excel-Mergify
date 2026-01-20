# Excel Mergify - Profesyonel Excel Birleştirme Uygulaması

## Proje Hakkında
Bu proje, iki farklı Excel dosyasını karşılaştırarak belirli kriterlere göre birleştiren modern bir ASP.NET Core web uygulamasıdır. Kullanıcıların teknik bilgiye ihtiyaç duymadan, sürükle-bırak yöntemiyle dosyalarını yükleyip sonuç alabilmesi için tasarlanmıştır.

## Nasıl Çalışır?
Uygulama iki temel dosya girdisi alır:

1.  **Ana Dosya (Kaynak):** İçinde verilerin aranacağı ana veri setidir. Bu dosyanın **G Sütunu (7. Sütun)**, referans dosya ile karşılaştırma yapmak için kullanılır.
2.  **Referans Dosya (Eşleşme):** Hangi verilerin filtreleneceğini belirleyen dosyadır. Bu dosyanın **A Sütunu (1. Sütun)** referans anahtarlarını barındırır.

### İşlem Adımları:
1.  Kullanıcı web arayüzü üzerinden her iki dosyayı yükler.
2.  Uygulama, Referans Dosyadaki A sütunu değerlerini hafızaya alır.
3.  Ana Dosyadaki her satırın G sütununu kontrol eder.
4.  Eğer G sütunundaki değer, referans listesinde varsa, o satırın ilk 12 sütunu (A-L) sonuç dosyasına eklenir.
5.  İşlem sonucunda oluşturulan "Kombine Excel Dosyası" otomatik olarak kullanıcının bilgisayarına indirilir.

## Teknolojiler
-   **Backend:** ASP.NET Core 6.0 MVC
-   **Excel İşlemleri:** EPPlus Kütüphanesi
-   **Frontend:** HTML5, CSS3 (Modern & Responsive Tasarım), Javascript
-   **Tasarım Dili:** Premium, temiz ve kullanıcı dostu arayüz.

## Kurulum ve Çalıştırma
Projeyi çalıştırmak için terminalde şu komutu kullanın:
`dotnet run`

Ardından tarayıcınızda `https://localhost:7010` adresine gidin.
