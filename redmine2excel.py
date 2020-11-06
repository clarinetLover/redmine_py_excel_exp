import sys, traceback, json, datetime
import openpyxl as xlsx
from redminelib import Redmine

# 設定ファイル読み込み
def import_settings():
    with open("setting.json") as settings:
        return json.load(settings)

# Redmineからチケット一覧を取得（コメント含む）
def get_issues():
    redmine_url = 'http://192.168.11.9/redmine'
    api_key = '257825921c793209eed344bb699aad027fc68a4e'
    project_id = 'sample_project'
    redmine = Redmine(redmine_url, key=api_key)
    issues = redmine.issue.filter(project_id=project_id)
    # 各チケットオブジェクトを取得し、チケットリストに追加
    issues_list = []
    for issue_tmp in issues:
        issues_list.append(redmine.issue.get(issue_tmp.id, include=['journals']))
    return issues_list

# Excelファイルを編集
def write_xlsx(settings, issue):
    # テンプレート読み込み
    try:
        wb = xlsx.load_workbook("test.xlsx")
        ws = wb["test"]
        for key in settings.keys():
            tmp = ""
            format_date = "{0:%Y/%m/%d}"
            if key == "journals":
                # コメントは記入日と記入者を含める
                for journal in issue[key]:
                    tmp += format_date.format(journal["created_on"]) + ' ' + journal["user"]["name"] + ' ' + journal["notes"] + "\n"
            elif isinstance(issue[key], datetime.date):
                tmp = format_date.format(issue[key])
            else:
                tmp = issue[key]
            ws[settings[key]].value = tmp
        wb.save("test" + str(issue["id"]) + ".xlsx")
    except OSError as os_err:
        tb = sys.exc_info()[2]
        traceback.print_tb(tb, limit=None, file=None)
    finally:
        wb.close()

# 主処理
# 設定ファイル読み込み
settings = import_settings()
issues = get_issues()
for issue in issues:
    write_xlsx(settings, issue)
