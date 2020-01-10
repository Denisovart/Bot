import vk_api
from vk_api.utils import get_random_id
from vk_api.bot_longpoll import VkBotLongPoll, VkBotEventType
import win32com.client as com_client

vk_session = vk_api.VkApi(token = "тут мой токен")
vk = vk_session.get_api()
longpoll = VkBotLongPoll(vk_session, "135775366")

def pod_sech(x, y):
	
	z = x/(y/10)

	return z

prov_a = False

for event in longpoll.listen():
	if event.type == VkBotEventType.MESSAGE_NEW: 
		if event.obj.text == "подбор сечения":
			if event.from_user:
				vk.messages.send(
						user_id = event.obj.from_id,
						random_id = get_random_id(),
						message = 'Напиши "Усилие: xxx" для ввода данных, усилие расчитывается в кН')
		elif event.obj.text[0:8] == "Усилие: ":
			a = event.obj.text[8:]
			if event.from_user:
				prov_a = True
				vk.messages.send(
						user_id = event.obj.from_id,
						random_id = get_random_id(),
						message = ' Напиши "Марка стали: ххх" для ввода данных. \nДанные по маркам стали: С235, С245, С255, С285, С345, С390, С440, С575')
		elif event.obj.text[0:13] == "Марка стали: ":
			if prov_a != True: 
				vk.messages.send(
						user_id = event.obj.from_id,
						random_id = get_random_id(),
						message = "Вы не ввели усилие")
			else:
				b = event.obj.text[13:]
				
				excel = com_client.Dispatch('Excel.Application')
				excel.Visible = False

				wb = excel.Workbooks.Open(r'C:/Users/User/bot/data.xlsx')
				ws = wb.Worksheets('zadacha')
				ws1 = wb.Worksheets('sortament')

				ws.Cells(4, 5).Value = b
				c = (ws.Cells(4, 7).Value)/10
				
				tp = pod_sech(float(a), float(c))

				prov = ws1.Cells(9, 17)
				ws1.Cells(3, 17).Value = tp
				ts = ws1.Cells(3, 19).Value

				if (prov is None) or (prov == 1) :
					new_tp = pod_sech(float(a), float(c))/2
					ws1.Cells(3, 17).Value = new_tp
					ts = 'Два двутавра с профилем ' + str(ws1.Cells(3, 19).Value)
					prov1 = ws1.Cells(4, 18)
					if int(prov1) == 38:
						ts = 'Не хватает данных, так же использование трёх двутавров чревато последствиями'

				wb.Close(SaveChanges = True)

				if event.from_user:
					vk.messages.send(
							user_id = event.obj.from_id,
							random_id = get_random_id(),
							message = "Твои данные: \nУсилие: " + a + " кН" + "\nРасчетное сопротивление стали: " + str(c) + " кН/см2" + "\nНомер требуемого двутавра: " + str(ts))
					prov_a = False
		elif event.obj.text == "4":
			if event.from_user:
				vk.messages.send(
						user_id = event.obj.from_id,
						random_id = get_random_id(),
						message = f)
		elif event.obj.text == "сортамент":
			if event.from_user:
				vk.messages.send(
						user_id = event.obj.from_id,
						random_id = get_random_id(),
						message = 'Напиши "сортамент: xxБх" для получения данных о двутавре с ГОСТ-26020-83')
		elif event.obj.text[0:11] == "сортамент: ":
			
			sort = event.obj.text[11:]
			
			excel = com_client.Dispatch('Excel.Application')
			excel.Visible = False
			
			wb = excel.Workbooks.Open(r'C:/Users/User/bot/data.xlsx')
			ws = wb.Worksheets('sortament')

			ws.Cells(7, 17).Value = sort
			check_sort = ws.Cells(9, 17).Value

			if (check_sort != 1) and (check_sort is not None):

				jx = ws.Cells(7, 19).Value
				wx = ws.Cells(8, 19).Value
				sx = ws.Cells(9, 19).Value
				ix = ws.Cells(10, 19).Value
				jy = ws.Cells(11, 19).Value
				wy = ws.Cells(12, 19).Value
				iy = ws.Cells(13, 19).Value

				wb.Close(SaveChanges = True)

				if event.from_user:
					vk.messages.send(
							user_id = event.obj.from_id,
							random_id = get_random_id(),
							message = 'Данные двутавра № ' + str(sort) + ":" + 
							"\nJx = " + str(jx) + " см4" + 
							"\nWx = " + str(wx) + " см3" + 
							"\nSx = " + str(sx) + " см3" + 
							"\nix = " + str(ix) + " см" + 
							"\nJy = " + str(jy) + " см4" + 
							"\nWy = " + str(wy) + " см3" + 
							"\niy = " + str(iy) + " см")
			else:
				wb.Close(SaveChanges = True)
				if event.from_user:
					vk.messages.send(
							user_id = event.obj.from_id,
							random_id = get_random_id(),
							message = 'Напиши корректнее "сортамент: xxБх"')
		elif event.obj.text == "сортамент кол двутавра":
			if event.from_user:
				vk.messages.send(
						user_id = event.obj.from_id,
						random_id = get_random_id(),
						message = 'Напиши "сортамент: xxКх" для получения данных о двутавре с ГОСТ-26020-83')
		elif event.obj.text[0:15] == "сортамент кдв: ":
			
			sort = event.obj.text[15:]
			
			excel = com_client.Dispatch('Excel.Application')
			excel.Visible = False
			
			wb = excel.Workbooks.Open(r'C:/Users/User/bot/data.xlsx')
			ws = wb.Worksheets('sortament_k')

			ws.Cells(7, 17).Value = sort
			check_sort = ws.Cells(9, 17).Value

			if (check_sort != 1) and (check_sort is not None):

				jx = ws.Cells(7, 19).Value
				wx = ws.Cells(8, 19).Value
				sx = ws.Cells(9, 19).Value
				ix = ws.Cells(10, 19).Value
				jy = ws.Cells(11, 19).Value
				wy = ws.Cells(12, 19).Value
				iy = ws.Cells(13, 19).Value

				wb.Close(SaveChanges = True)

				if event.from_user:
					vk.messages.send(
							user_id = event.obj.from_id,
							random_id = get_random_id(),
							message = 'Данные двутавра № ' + str(sort) + ":" + 
							"\nJx = " + str(jx) + " см4" + 
							"\nWx = " + str(wx) + " см3" + 
							"\nSx = " + str(sx) + " см3" + 
							"\nix = " + str(ix) + " см" + 
							"\nJy = " + str(jy) + " см4" + 
							"\nWy = " + str(wy) + " см3" + 
							"\niy = " + str(iy) + " см")
			else:
				wb.Close(SaveChanges = True)
				if event.from_user:
					vk.messages.send(
							user_id = event.obj.from_id,
							random_id = get_random_id(),
							message = 'Напиши корректнее "сортамент: xxБх"')
		elif event.obj.text == "help":
			if event.from_user:
				vk.messages.send(
						user_id = event.obj.from_id,
						random_id = get_random_id(),
						message = "Я бот созданный Денисовым Артёмом для подбора сечения балки и просмотра данных о двутавре. \nМои команды: \n1 - Напиши 'подбор сечения' для расчета  \n2 - Напиши 'сортамент' для просмотра данных о двутавре \n3 - Напиши 'сортамент кол двутавра' для просмотра данных о колонном двутавре")
		else:
			if event.from_user:
				vk.messages.send(
						user_id = event.obj.from_id,
						random_id = get_random_id(),
						message = 'Я вас не понимаю, для просмотра моих возможностей напишите "help"')






