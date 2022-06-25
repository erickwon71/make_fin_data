# make_fin_data
financial data collecting and handling

resource : 재무제표 statement별 input, corp code xml
    sc : 현금흐름표
    sf : 재무상태표
    si : 손익계산서
    CORPCODE.xml
findata : 재무 data 엑셀 파일 모아놓은 폴더
tests : unit test
docs : 문서 file들
tmp : 임시 저장 file
main.py : main
make_flist.py : findata에 있는 file 중 기업명에 해당하는 file list를 가져오는 module
dart_interface.py : opendart interface를 위한 module
dl_finance_xls.py : dart에서 재무 엑셀 파일 download
