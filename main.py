import pandas as pd
import win32com.client as win32

outlook = win32.Dispatch('outlook.application')

mail = outlook.CreateItem(0)
mail.To = 'prof.kaiolobato@gmail.com'
mail.Subject = 'E-mail teste PYTHON'
mail.Body = '''Olá... Isso é um teste!'''
mail.Send()