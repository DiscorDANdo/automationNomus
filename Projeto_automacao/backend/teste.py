from datetime import date, timedelta

data = date.today()
data_futura = data + timedelta(days=10)
data_br = data_futura.strftime("%d/%m/%Y")

print(data_br)