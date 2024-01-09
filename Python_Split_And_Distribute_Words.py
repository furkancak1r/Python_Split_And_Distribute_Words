# İsim soyisim birleşik yazılan hücreleri ayrı kolonlara yazar. Örneğin "Furkan Çakır" yazan hücreyi "Furkan" ve "Çakır" olarak iki hücreye böler.
import openpyxl

# Excel dosyasını aç
wb = openpyxl.load_workbook(r"C:\Users\furkan.cakir\Desktop\FurkanPRS\Kodlar\Finans & Muhasebe\exceller\Cariler\OCPR-İlgili Kişiler.xlsx")
sheet = wb['INKOOL']  # Sayfa adını kullan

# 5. satırdan 7581. satıra kadar dolaş
for row in range(5, 7582):
    cell_value = sheet[f'D{row}'].value

    # Eğer hücre boş değilse, kelimelere ayır
    if cell_value:
        words = cell_value.split()

        # Eğer iki kelime varsa, birinci kelimeyi E'ye, ikinci kelimeyi G'ye yaz
        if len(words) == 2:
            sheet[f'E{row}'].value = words[0]
            sheet[f'G{row}'].value = words[1]
        else:
            # İlk kelimeyi E sütununa yaz
            sheet[f'E{row}'].value = words[0] if len(words) > 0 else ''

            # İkinci kelimeyi F sütununa yaz
            sheet[f'F{row}'].value = words[1] if len(words) > 1 else ''

            # Üçüncü ve sonraki kelimeleri G sütununa birleştirip yaz
            if len(words) > 2:
                sheet[f'G{row}'].value = ' '.join(words[2:])
            else:
                sheet[f'G{row}'].value = ''

# Excel dosyasını kaydet
wb.save(r"C:\Users\furkan.cakir\Desktop\FurkanPRS\Kodlar\Finans & Muhasebe\exceller\Cariler\OCPR-İlgili Kişiler.xlsx")
