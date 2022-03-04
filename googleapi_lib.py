from googleapiclient.discovery import build
from google.oauth2 import service_account
import json


class google_sheet:
	""" Classe qui facilite l'interraction avec l'api google sheet
	spreadsheet_id : id de la feuille de calcul, présent dans l'url 
	key_file : fichier de clé obtenu lors de l'ajout d'un utilisteur de test au projet google
	scopes : liste qui contient https://www.googleapis.com/auth/spreadsheets pour les google sheet
	"""

	def __init__(self, spreadsheet_id, key_file, scopes=['https://www.googleapis.com/auth/spreadsheets']):
		self.credentials = service_account.Credentials.from_service_account_file(key_file, scopes=scopes)
		self.sheet = build('sheets', 'v4', credentials=self.credentials).spreadsheets()
		self.spreadsheet_id = spreadsheet_id

		# Récupération des sheetId
		list_id = self.sheet.get(spreadsheetId=self.spreadsheet_id).execute()
		self.sheetId = {}
		for item in list_id["sheets"]:
			self.sheetId[item["properties"]["title"]] = item["properties"]["sheetId"]

	def read(self, range):
		""" Lire un range de cellules 
		range doit être une liste correspondant à la taille du range """
		return self.sheet.values().get(spreadsheetId=self.spreadsheet_id,range=range).execute().get('values', [])

	def write(self, range, data):
		""" Ecrire des données dans une plage de cellules """
		body = {'values': data}
		result = self.sheet.values().update(spreadsheetId=self.spreadsheet_id, range=range, valueInputOption="RAW", body=body).execute() #  value_input_option,
		return f"{result.get('updatedCells')} cells updated."

	def auto_fit(self, sheet_name=0, column_range=None, row_range=None):
		""" Taille automatique des colonnes ou des lignes
		sheet_name : nom de la feuille sur laquelle travailler. S'il n'y en a qu'une ou si on travaille sur la première, ce pramètre est facultatif
		column_range et row_range sont des listes ou tuples contenant le début et la fin du range à modifier
		"""
		if sheet_name != 0:
			sheet_name = self.sheetId[sheet_name]

		request_body = {"requests":[]}


		if column_range is not None:
			request_body["requests"].append({
				"autoResizeDimensions": {
					"dimensions": {
						"sheetId": sheet_name,
						"dimension": "COLUMNS",
						"startIndex": column_range[0],
						"endIndex": column_range[1]
					}
				}
			})
		if row_range is not None:
			request_body["requests"].append({
				"autoResizeDimensions": {
					"dimensions": {
						"sheetId": sheet_name,
						"dimension": "ROWS",
						"startIndex": row_range[0],
						"endIndex": row_range[1]
					}
				}
			})
		
		return self.sheet.batchUpdate(spreadsheetId=self.spreadsheet_id, body=request_body).execute()