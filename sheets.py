import os
import json
from flask import Blueprint, request, jsonify
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

sheets_bp = Blueprint('sheets', __name__)

# Google Sheets APIのスコープ
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

def get_sheets_service():
    """Google Sheets APIサービスを取得"""
    try:
        # 認証情報ファイルのパス
        # credentials_path = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(__file__))), 'credentials.json')
      
        # if not os.path.exists(credentials_path):
         #    raise FileNotFoundError(f"認証情報ファイルが見つかりません: {credentials_path}")
    

        def get_sheets_service():
            """Google Sheets APIサービスを取得"""
            try:
                # 環境変数から認証情報を取得
                credentials_json = os.environ.get("GOOGLE_APPLICATION_CREDENTIALS_JSON")
                if not credentials_json:
                    raise ValueError("GOOGLE_APPLICATION_CREDENTIALS_JSON 環境変数が設定されていません")
                
                credentials_info = json.loads(credentials_json)
                credentials = service_account.Credentials.from_service_account_info(
                    credentials_info, scopes=SCOPES
                )
                
                service = build("sheets", "v4", credentials=credentials)
                return service
            except Exception as e:
                print(f"Google Sheets API認証エラー: {e}")
                raise


     
        # サービスアカウント認証
        credentials = service_account.Credentials.from_service_account_file(
            credentials_path, scopes=SCOPES
        )
        
        # Google Sheets APIサービスを構築
        service = build('sheets', 'v4', credentials=credentials)
        return service
    except Exception as e:
        print(f"Google Sheets API認証エラー: {e}")
        raise

@sheets_bp.route('/write-to-sheets', methods=['POST'])
def write_to_sheets():
    """CSVデータをGoogle Sheetsに書き込む"""
    try:
        # リクエストデータを取得
        data = request.get_json()
        
        if not data:
            return jsonify({'error': 'リクエストデータが空です'}), 400
        
        spreadsheet_id = data.get('spreadsheet_id')
        sheet_name = data.get('sheet_name', 'Sheet1')
        csv_data = data.get('csv_data')
        clear_existing = data.get('clear_existing', True)
        
        if not spreadsheet_id:
            return jsonify({'error': 'spreadsheet_idが必要です'}), 400
        
        if not csv_data:
            return jsonify({'error': 'csv_dataが必要です'}), 400
        
        # Google Sheets APIサービスを取得
        service = get_sheets_service()
        sheet = service.spreadsheets()
        
        # 既存データをクリア（オプション）
        if clear_existing:
            clear_range = f"{sheet_name}!A:Z"
            sheet.values().clear(
                spreadsheetId=spreadsheet_id,
                range=clear_range
            ).execute()
        
        # CSVデータをGoogle Sheetsに書き込み
        range_name = f"{sheet_name}!A1"
        
        # CSVデータを2次元配列に変換
        if isinstance(csv_data[0], dict):
            # 辞書形式の場合、ヘッダーとデータに分離
            headers = list(csv_data[0].keys())
            values = [headers]
            for row in csv_data:
                values.append([str(row.get(header, '')) for header in headers])
        else:
            # すでに2次元配列の場合
            values = csv_data
        
        body = {
            'values': values
        }
        
        result = sheet.values().update(
            spreadsheetId=spreadsheet_id,
            range=range_name,
            valueInputOption='RAW',
            body=body
        ).execute()
        
        updated_cells = result.get('updatedCells', 0)
        
        return jsonify({
            'success': True,
            'message': f'{updated_cells}個のセルが更新されました',
            'updated_cells': updated_cells,
            'spreadsheet_url': f'https://docs.google.com/spreadsheets/d/{spreadsheet_id}'
        })
        
    except HttpError as error:
        print(f"Google Sheets API HTTPエラー: {error}")
        return jsonify({
            'error': f'Google Sheets APIエラー: {error}',
            'details': str(error)
        }), 500
        
    except FileNotFoundError as error:
        print(f"認証ファイルエラー: {error}")
        return jsonify({
            'error': '認証情報ファイルが見つかりません',
            'details': str(error)
        }), 500
        
    except Exception as error:
        print(f"予期しないエラー: {error}")
        return jsonify({
            'error': '予期しないエラーが発生しました',
            'details': str(error)
        }), 500

@sheets_bp.route('/test-connection', methods=['GET'])
def test_connection():
    """Google Sheets API接続テスト"""
    try:
        service = get_sheets_service()
        
        # テスト用のスプレッドシート情報を取得
        test_spreadsheet_id = request.args.get('spreadsheet_id')
        
        if not test_spreadsheet_id:
            return jsonify({
                'success': True,
                'message': 'Google Sheets API認証は正常です',
                'note': 'spreadsheet_idパラメータを指定すると、特定のスプレッドシートへのアクセステストができます'
            })
        
        # 指定されたスプレッドシートの情報を取得
        sheet = service.spreadsheets()
        spreadsheet = sheet.get(spreadsheetId=test_spreadsheet_id).execute()
        
        title = spreadsheet.get('properties', {}).get('title', 'Unknown')
        sheet_count = len(spreadsheet.get('sheets', []))
        
        return jsonify({
            'success': True,
            'message': 'Google Sheets API接続テスト成功',
            'spreadsheet_title': title,
            'sheet_count': sheet_count,
            'spreadsheet_url': f'https://docs.google.com/spreadsheets/d/{test_spreadsheet_id}'
        })
        
    except HttpError as error:
        return jsonify({
            'success': False,
            'error': f'Google Sheets APIエラー: {error}',
            'details': str(error)
        }), 500
        
    except Exception as error:
        return jsonify({
            'success': False,
            'error': '接続テストに失敗しました',
            'details': str(error)
        }), 500

@sheets_bp.route('/get-sheet-info', methods=['GET'])
def get_sheet_info():
    """スプレッドシートの情報を取得"""
    try:
        spreadsheet_id = request.args.get('spreadsheet_id')
        
        if not spreadsheet_id:
            return jsonify({'error': 'spreadsheet_idパラメータが必要です'}), 400
        
        service = get_sheets_service()
        sheet = service.spreadsheets()
        
        # スプレッドシート情報を取得
        spreadsheet = sheet.get(spreadsheetId=spreadsheet_id).execute()
        
        title = spreadsheet.get('properties', {}).get('title', 'Unknown')
        sheets_info = []
        
        for sheet_data in spreadsheet.get('sheets', []):
            sheet_properties = sheet_data.get('properties', {})
            sheets_info.append({
                'sheet_id': sheet_properties.get('sheetId'),
                'title': sheet_properties.get('title'),
                'index': sheet_properties.get('index'),
                'sheet_type': sheet_properties.get('sheetType', 'GRID')
            })
        
        return jsonify({
            'success': True,
            'spreadsheet_title': title,
            'spreadsheet_id': spreadsheet_id,
            'sheets': sheets_info,
            'spreadsheet_url': f'https://docs.google.com/spreadsheets/d/{spreadsheet_id}'
        })
        
    except HttpError as error:
        return jsonify({
            'success': False,
            'error': f'Google Sheets APIエラー: {error}',
            'details': str(error)
        }), 500
        
    except Exception as error:
        return jsonify({
            'success': False,
            'error': 'スプレッドシート情報の取得に失敗しました',
            'details': str(error)
        }), 500

