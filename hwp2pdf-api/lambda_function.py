import json
import boto3
import os

# S3 클라이언트 초기화
s3 = boto3.client('s3')
BUCKET_NAME = 'hwp2pdf-storage-kstlabs'  # 앞서 생성한 버킷 이름으로 변경

def lambda_handler(event, context):
    """
    Lambda 함수 핸들러
    """
    # 기본 응답
    return {
        'statusCode': 200,
        'body': json.dumps({
            'message': 'Hello from Lambda!'
        })
    }