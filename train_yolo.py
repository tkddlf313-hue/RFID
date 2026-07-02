# -*- coding: utf-8 -*-
"""
============================================================
 스마트 팩토리용 YOLO 커스텀 자재 학습 템플릿
 ============================================================
 이 스크립트는 직접 촬영하고 라벨링한 자재 데이터를 사용하여 
 나만의 YOLO 모델을 학습시키는 과정을 안내하는 템플릿입니다.

 [사전 준비 사항]
 1. 데이터 수집: 자재 사진을 다양한 각도에서 최소 100장 이상 촬영합니다.
 2. 라벨링: Roboflow(https://roboflow.com/) 등의 툴을 이용하여 
    각 이미지 내 자재의 경계 상자(Bounding Box)를 지정하고 YOLO 포맷으로 다운로드합니다.
 3. 폴더 구조 구성:
    dataset/
      ├── data.yaml        # 클래스 수, 이름, 경로 정의 파일
      ├── train/
      │     ├── images/    # 학습용 이미지들 (.jpg)
      │     └── labels/    # 학습용 라벨 텍스트 (.txt)
      └── val/
            ├── images/    # 검증용 이미지들
            └── labels/    # 검증용 라벨 텍스트
"""

import os
from ultralytics import YOLO

def check_dataset_ready(yaml_path):
    """데이터셋 구성 파일(data.yaml)의 존재 여부를 검사합니다."""
    if not os.path.exists(yaml_path):
        print(f"[오류] '{yaml_path}' 파일이 존재하지 않습니다.")
        print("데이터셋 라벨링 완료 후 다운로드 받은 data.yaml 파일 경로를 확인해 주세요.")
        return False
    return True

def start_training(yaml_path="dataset/data.yaml", base_model="yolo11n.pt", epochs=50, imgsz=640):
    """
    YOLO 커스텀 모델 학습을 실행합니다.
    - yaml_path: 데이터셋 설정 파일 경로
    - base_model: 사전 학습된 베이스 모델 (yolo11n, yolo11s 등)
    - epochs: 학습을 반복할 횟수 (테스트용은 10~50회, 실전용은 100~300회 권장)
    - imgsz: 이미지 크기 (기본값 640)
    """
    if not check_dataset_ready(yaml_path):
        return

    print("====================================================")
    print(f" YOLO 커스텀 모델 학습 시작")
    print(f" - 베이스 모델: {base_model}")
    print(f" - 학습 반복수: {epochs} epochs")
    print(f" - 이미지 크기: {imgsz}")
    print("====================================================")

    # 1. 사전 학습된 기본 모델 로드 (가중치 다운로드)
    model = YOLO(base_model)

    # 2. 모델 학습 시작 (학습 결과는 runs/detect/train/ 폴더 아래 저장됩니다)
    model.train(
        data=yaml_path,
        epochs=epochs,
        imgsz=imgsz,
        device="cpu",      # GPU가 장착된 PC라면 '0' 또는 'cuda' 입력 권장 (CPU는 다소 느릴 수 있음)
        workers=2,         # 데이터 로더 스레드 수
        verbose=True
    )

    print("\n[학습 완료] runs/detect/train/weights/best.pt 파일이 생성되었습니다.")
    print("해당 파일을 복사하여 대시보드의 './models/' 폴더에 넣고 활용하세요!")

if __name__ == "__main__":
    # 실행 예시 (실제 실행하려면 아래 주석을 해제하고 yaml 경로를 확인하세요)
    # start_training(yaml_path="./dataset/data.yaml", base_model="yolo11n.pt", epochs=50)
    
    print("YOLO 학습 스크립트 로드 완료.")
    print("학습을 시작하려면 이 파일의 하단 'start_training' 함수의 주석을 해제하고")
    print("`python train_yolo.py` 명령어로 실행하세요.")
