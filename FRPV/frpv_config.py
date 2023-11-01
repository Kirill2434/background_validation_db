import glob
from pathlib import PureWindowsPath

main_folder = PureWindowsPath(r'C:\source_data')

R_1 = list(range(1, 27))
R_2 = list(range(1, 18))
R_3 = list(range(1, 18))

SHEET_NAME_FRPV = ['Раздел 1', 'Раздел 2', 'Раздел 3']
COL_NAME_FRPV = [R_1, R_2, R_3]

FRPV_CHECK_REPORT = r'C:\generation_results\check_report_file.xlsx'
FRPV_CHECK_REPORT_N = r'C:\generation_results\check_report_file'

all_frpv_files = glob.glob(r'C:\source_data\*.xlsx')
