# Documentation : https://developers.google.com/sheets/api/samples/sheet

from googleapiclient.discovery import build # Installer avec 'pip install google-api-python-client'
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
		range doit être une liste correspondant à la taille du range 
		Il peut aussi être une str par ex. "A1:B5"
		"""
		return self.sheet.values().get(spreadsheetId=self.spreadsheet_id,range=range).execute().get('values', [])

	def write(self, range, data):
		""" Ecrire des données dans une plage de cellules 
			Les données à écrire doivent être une liste ou un tuple à 2 dimentions même s'il n'y a qu'une ligne à écrire
		"""
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

	def color(self, column_range, row_range, color, sheet_name=0):
		""" Coloriser un range donné
			column_range doit être une liste de 2 éléments contenant la première puis la dernière colonne à colorer
			row_range doit être une liste de 2 éléments contenant la première puis la dernière ligne à colorer
			le start index doit être de la 1e ligne -1
			Mettre à None si inutilisé
			color doit être une liste de 4 éléments contenant les 3 couleurs respectivement (R,G,B, A) valeurs de 0 à 255
		"""
		if sheet_name != 0:
			sheet_name = self.sheetId[sheet_name]

		# Conversion des couleurs 0-255 à 0-1
		converted_colors = [item/255 for item in color]


		request_body = {
			"requests":[
				{
			      "repeatCell": 
			      {
			        "range": 
			        {
			          "sheetId": sheet_name,
			          "startRowIndex": row_range[0],
			          "endRowIndex": row_range[1],
			          "startColumnIndex": column_range[0],
			          "endColumnIndex": column_range[1]
			        },
			        "cell": 
			    	{
			    	"userEnteredFormat"	: 
			    		{
			    			 "backgroundColor": 
			                  {
			                    "red": converted_colors[0],
			                    "green": converted_colors[1],
			                    "blue": converted_colors[2],
			                    "alpha": converted_colors[3],
			                  }
			    		}
			    	},
			    	"fields":"userEnteredFormat.backgroundColor",
			       
			      }
			    }
			  ]
			}

		return self.sheet.batchUpdate(spreadsheetId=self.spreadsheet_id, body=request_body).execute()

	def new_sheet(self, name):
		""" Création d'un nouvelle feuille """
		request_body = {"requests":[{"addSheet":{'properties': {'title': name}}}]}
		self.spreadsheet.batchUpdate(spreadsheetId=self.spreadsheet_id,body=request_body).execute()
		self.read_sheets_id()

	def delete_sheet(self, sheet_id):
		""" Suppression d'une feuille """
		request_body = {"requests":[{"deleteSheet":{'sheetId': sheet_id}}]}
		self.spreadsheet.batchUpdate(spreadsheetId=self.spreadsheet_id,body=request_body).execute()

	def hide_column(self,sheet_name, range):
		""" cacher une colonne """

		sheet_id = self.sheets[sheet_name]

		requests = [{
			'updateDimensionProperties': {
			"range": {
			  "sheetId": sheet_id,
			  "dimension": 'COLUMNS',
			  "startIndex": range[0],
			  "endIndex": range[1],
			},
			"properties": {
			  "hiddenByUser": True,
			},
			"fields": 'hiddenByUser',
			}}]

		body = {'requests': requests}

		return self.spreadsheet.batchUpdate(spreadsheetId=self.spreadsheet_id,body=body).execute()