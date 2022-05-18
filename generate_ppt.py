import json
import io
import boto3
import re

from pptx import Presentation
from pptx.enum.dml import MSO_COLOR_TYPE
from pptx.dml.color import RGBColor

def lambda_handler(event, context):
  s3 = boto3.client('s3')
  s3_response_object = s3.get_object(Bucket='bucket-name', Key=event['key'])
  object_content = s3_response_object['Body'].read()

  prs = Presentation(io.BytesIO(object_content))

 #UPDATE ACCORDING TO YOUR WORK

  file_key = event['title'] + '.pptx'

  with io.BytesIO() as out:
      prs.save(out)
      out.seek(0)
      s3.upload_fileobj(out, 'bucket-name', file_key)

  file_url = s3.generate_presigned_url('get_object',
                                                Params={'Bucket': 'bucket-name','Key': file_key},
                                                ExpiresIn=3600)
  return {
    'statusCode': 200,
    'body': json.dumps(file_url)
  }
