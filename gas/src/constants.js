const SHEET_ID = '1hnFUAN514puTxfHNqkZZiKdmT7sN8B2EkGGBGs_FUjQ';
const SPREADSHEET_URL =
  'https://docs.google.com/spreadsheets/d/1hnFUAN514puTxfHNqkZZiKdmT7sN8B2EkGGBGs_FUjQ';
const SUPABASE_STORAGE_BUCKET = 'inquiry-files';

const SHEET_NAMES = {
  LEDGER: '案件台帳',
  INTERNAL_LOG: '内部ログ',
  SIMILAR_CASES: '類似案件サンプル',
};

const COLUMNS = {
  INTERNAL_LOG: {
    ID: 'ID',
    CREATED_AT: '作成日時',
    SOURCE: '受付経路',
    STATUS: 'ステータス',
    CUSTOMER_NAME: '顧客名',
    CONTACT_NAME: '担当者名',
    PROJECT_NAME: '案件名',
    DRAWING_NUMBER: '図面番号',
    MATERIAL: '材質',
    SIZE: 'サイズ',
    QUANTITY: '数量',
    DUE_DATE: '希望納期',
    NOTES: '備考',
    RAW_TEXT: '元テキスト',
    LINE_USER_ID: 'LINEユーザーID',
    AI_EXTRACTED_JSON: 'AI抽出結果JSON',
    SIMILAR_CASE: '類似案件候補',
    PAST_UNIT_PRICE: '過去単価',
    SUGGESTED_PRICE: '今回提示単価',
    SPREADSHEET_UPDATED_AT: 'スプレッドシート更新日時',
  },

  SIMILAR_CASES: {
    PROJECT_ID: '案件ID',
    RECEIVED_AT: '受付日時',
    CUSTOMER_NAME: '顧客名',
    CONTACT_NAME: '担当者名',
    PROJECT_NAME: '案件名',
    DRAWING_NUMBER: '図面番号',
    PROCESS_TYPE: '加工分類',
    MATERIAL: '材質',
    SIZE: 'サイズ',
    QUANTITY: '数量',
    DUE_DATE: '希望納期',
    PAST_UNIT_PRICE: '過去単価',
    NOTES: '備考',
  },

  LEDGER: {
    RECEIVED_AT: '受付日時',
    STATUS: 'ステータス',
    SOURCE: '受付経路',
    CUSTOMER_NAME: '顧客名',
    CONTACT_NAME: '担当者名',
    EMAIL: 'メールアドレス',
    PHONE: '電話番号',
    INQUIRY: '案件名',
    DUE_DATE: '希望納期',
    MATERIAL: '材質',
    SIZE_THICKNESS: 'サイズ・板厚',
    QUANTITY: '数量',
    NOTES: '補足事項',
    DRAWING_URL: '図面URL',
    RAW_JSON: 'raw_json',
  },
};

const PROPS = PropertiesService.getScriptProperties();
const GEMINI_API_KEY = PROPS.getProperty('GEMINI_API_KEY');
const GEMINI_MODEL = PROPS.getProperty('GEMINI_MODEL');
const SUPABASE_URL = PROPS.getProperty('SUPABASE_URL');
const SUPABASE_KEY = PROPS.getProperty('SUPABASE_KEY');
const LINE_CHANNEL_ACCESS_TOKEN = PROPS.getProperty(
  'LINE_CHANNEL_ACCESS_TOKEN'
);
const LINE_NOTIFY_USER_ID = PROPS.getProperty('LINE_NOTIFY_USER_ID');
const LINE_RICH_MENU_IMAGE_FILE_ID = PROPS.getProperty(
  'LINE_RICH_MENU_IMAGE_FILE_ID'
);
