import pandas as pd
import openpyxl
import os
import numpy as np
import re
import json


def transform_excel_file(input_file="input.xlsx", output_file="output.xlsx", exception_file="exceptions.json"):
    """
    1번 Excel 파일을 2번 파일과 같은 형식으로 변환하는 함수
    
    Args:
        input_file (str): 입력 파일 경로 (기본값: input.xlsx)
        output_file (str): 출력 파일 경로 (기본값: output.xlsx)
        exception_file (str): 예외 상품 목록이 담긴 JSON 파일 경로 (기본값: exceptions.json)
    """
    try:
        print(f"파일 로딩 중: {input_file}")
        df = pd.read_excel(input_file, sheet_name=0, dtype=str)
        
        # 예외 목록 로드
        exception_products = []
        if os.path.exists(exception_file):
            try:
                with open(exception_file, 'r', encoding='utf-8') as f:
                    exception_data = json.load(f)
                    if "exception_products" in exception_data:
                        exception_products = exception_data["exception_products"]
                print(f"예외 상품 목록 로드 완료: {len(exception_products)}개 항목")
            except Exception as e:
                print(f"예외 목록 로드 중 오류 발생: {str(e)}")
                print("예외 없이 계속 진행합니다.")
        else:
            print(f"예외 파일 '{exception_file}'이 존재하지 않습니다. 예외 없이 계속 진행합니다.")
        
        required_columns = ['주문번호', '상태', '상품명-옵션명', '관리용상품명', '수량', 
                           '받는분', '받는분 연락처', '배송지 우편번호', '도로명 주소', '배송메시지']
        
        missing_columns = [col for col in required_columns if col not in df.columns]
        if missing_columns:
            print(f"오류: 필요한 열이 없습니다: {', '.join(missing_columns)}")
            return False
        
        # NaN 값 처리
        df = df.replace({np.nan: ''})
        
        # 1. 데이터 전처리 - 필요한 열만 선택
        required_data = df[[
            '주문번호', '상태', '상품명-옵션명', '관리용상품명', '수량', 
            '받는분', '받는분 연락처', '배송지 우편번호', '도로명 주소', '배송메시지'
        ]].copy()
        
        # 빈 행 제거
        required_data = required_data[required_data['주문번호'] != '']
        
        # 2. 고객별 주문 데이터 그룹화
        customers = {}
        customer_order_count = {}  # 고객별 주문 카운터 (예외 상품용)
        
        for _, row in required_data.iterrows():
            # 상품명 가져오기
            product_name = str(row['관리용상품명']).strip()
            
            # 고객 키 생성 (받는분, 연락처, 우편번호, 주소, 배송메시지 조합)
            customer_base_key = f"{row['받는분']}_{row['받는분 연락처']}_{row['배송지 우편번호']}_{row['도로명 주소']}_{row['배송메시지']}"
            
            # 예외 상품인지 확인
            is_exception = product_name in exception_products
            
            # 예외 상품인 경우 고유한 키를 생성하여 별도 행으로 처리
            if is_exception:
                if customer_base_key not in customer_order_count:
                    customer_order_count[customer_base_key] = 0
                customer_order_count[customer_base_key] += 1
                customer_key = f"{customer_base_key}_{customer_order_count[customer_base_key]}"
            else:
                customer_key = customer_base_key
            
            # 수량 추출
            quantity_str = str(row['수량']).strip()
            quantity = int(quantity_str) if quantity_str and quantity_str.isdigit() else 1
            
            # 장수 표현이 있는지 체크 (예: "상품명 1000장")
            match = re.search(r'(.+?)\s+(\d+)장$', product_name)
            if match:
                # 상품 기본명과 장수 분리
                base_name = match.group(1).strip()
                original_quantity = int(match.group(2))
                # 상품명은 기본명만 사용하고, 수량은 원래 장수 × 주문 수량
                product_name = base_name
                total_quantity = original_quantity * quantity
            else:
                # 장수 표현이 없는 경우 수량은 그대로
                total_quantity = quantity
            
            # 고객 정보 저장 또는 업데이트
            if customer_key not in customers:
                customers[customer_key] = {
                    'name': row['받는분'],
                    'phone': row['받는분 연락처'],
                    'postal_code': row['배송지 우편번호'],
                    'address': row['도로명 주소'],
                    'message': row['배송메시지'],
                    'products': {},
                    'is_exception': is_exception
                }
            
            # 상품별로 저장 (동일한 제품이면 수량 합산, 예외 상품은 항상 새로 추가)
            if product_name in customers[customer_key]['products'] and not is_exception:
                customers[customer_key]['products'][product_name] += total_quantity
            else:
                customers[customer_key]['products'][product_name] = total_quantity
        
        # 3. 변환된 데이터 생성
        transformed_rows = []
        
        for customer_info in customers.values():
            # 상품 정보를 형식화하여 리스트로 변환
            product_list = []
            for product_name, count in customer_info['products'].items():
                if "장" in product_name:
                    # 이미 "장"이 포함된 상품명은 그대로 사용
                    product_str = product_name
                else:
                    # "장"이 없는 상품명에 수량 추가 (수량이 1이 아닐 경우만)
                    if count > 1:
                        product_str = f"{product_name} {count}장"
                    else:
                        product_str = product_name
                product_list.append(product_str)
            
            # 상품이 여러 개인 경우 콤마로 구분하여 병합
            products_str = " ,".join(product_list)
            
            # 2번 파일 형식에 맞게 데이터 구성
            transformed_row = [
                customer_info['name'],         # A: 받는분 
                customer_info['phone'],        # B: 받는분 연락처
                customer_info['postal_code'],  # C: 배송지 우편번호
                customer_info['address'],      # D: 도로명 주소
                customer_info['message'],      # E: 배송메시지 
                products_str,                  # F: 상품명
                "1"                            # G: 수량
            ]
            
            transformed_rows.append(transformed_row)
        
        # 데이터가 없는 경우 처리
        if not transformed_rows:
            print("변환할 데이터가 없습니다.")
            return False
            
        # 새 워크북 생성
        new_workbook = openpyxl.Workbook()
        new_sheet = new_workbook.active
        new_sheet.title = "주문관리목록"  # 첫 번째 시트 이름 설정
        
        # 변환된 데이터 입력 - 모든 셀을 문자열로 저장
        for row_idx, row_data in enumerate(transformed_rows, 1):
            for col_idx, cell_value in enumerate(row_data, 1):
                cell = new_sheet.cell(row=row_idx, column=col_idx)
                cell.value = cell_value
                
                # 전화번호(B열)는 문자열 서식 적용
                if col_idx == 2:  # B열 (전화번호)
                    cell.number_format = '@'  # 문자열 형식 지정
        
        # 원본 워크북의 다른 시트 복사
        original_workbook = openpyxl.load_workbook(input_file)
        original_sheets = original_workbook.sheetnames
        
        for sheet_name in original_sheets[1:]:  # 첫 번째 시트를 제외한 나머지 시트 복사
            source_sheet = original_workbook[sheet_name]
            target_sheet = new_workbook.create_sheet(title=sheet_name)
            
            # 셀 복사
            for row in source_sheet.rows:
                for cell in row:
                    target_sheet[cell.coordinate] = cell.value
        
        # 결과 파일 저장
        try:
            new_workbook.save(output_file)
            print(f"파일 변환 완료: {output_file}")
            return True
        except Exception as e:
            print(f"파일 저장 중 오류 발생: {str(e)}")
            return False
            
    except Exception as e:
        print(f"파일 변환 중 오류 발생: {str(e)}")
        import traceback
        traceback.print_exc()
        return False


# 예외 목록을 생성하는 함수
def create_exception_list(exception_list, output_file="exceptions.json"):
    """
    예외 상품명 목록을 JSON 파일로 저장하는 함수
    
    Args:
        exception_list (list): 예외 처리할 상품명 목록
        output_file (str): 출력할 JSON 파일 경로 (기본값: exceptions.json)
    """
    try:
        data = {"exception_products": exception_list}
        with open(output_file, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        print(f"예외 목록 파일 생성 완료: {output_file}")
        return True
    except Exception as e:
        print(f"예외 목록 파일 생성 중 오류 발생: {str(e)}")
        return False


if __name__ == "__main__":
    # 입력 파일 존재 확인
    if not os.path.exists("input.xlsx"):
        print("오류: 'input.xlsx' 파일이 현재 폴더에 존재하지 않습니다.")
    else:
        transform_excel_file()