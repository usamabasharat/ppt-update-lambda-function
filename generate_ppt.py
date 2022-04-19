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

  trans_num = ['First', 'Second', 'Third', 'Fourth', 'Fifth', 'Sixth', 'Seventh', 'Eighth', 'Ninth', 'Tenth'];
  print(event)
  for slide in prs.slides:
    for shape in slide.shapes:
      if not shape.has_text_frame:
        continue
      for paragraph in shape.text_frame.paragraphs:
        if re.search('%(.*)%', paragraph.text):
          text = re.search('%(.*)%', paragraph.text).group(1)
          if paragraph.runs[0].font.color.type == MSO_COLOR_TYPE.RGB:
            color = paragraph.runs[0].font.color.rgb
          else:
            color = None
          size = paragraph.runs[0].font.size
          bold = paragraph.runs[0].font.bold
          font_name = paragraph.runs[0].font.name

          if "Insight" in text:
            for insight in range(len(event['insights'])):
              if text == f"Insight {insight + 1}":
                paragraph.text = paragraph.text.replace(text, event['insights_title'][insight] if event['insights_title'][insight] else 'Insight: N/A')
                paragraph.text = paragraph.text.replace('%', '')
                if color:
                  paragraph.runs[0].font.color.rgb = color
                paragraph.runs[0].font.size = size
                paragraph.runs[0].font.bold = bold
                paragraph.runs[0].font.name = font_name

            for x in range(len(event['insights_data_points'])):
              for y in range(len(event['insights_data_points'][x])):
                if text == f'Insight {x + 1} Data Point {y + 1}':
                  print('DP')
                  paragraph.text = paragraph.text.replace(text, event['insights_data_points'][x][y] if event['insights_data_points'][x][y] else 'Data Point: N/A')
                  paragraph.text = paragraph.text.replace('%', '')
                  if color:
                    paragraph.runs[0].font.color.rgb = color
                  paragraph.runs[0].font.size = size
                  paragraph.runs[0].font.bold = bold
                  paragraph.runs[0].font.name = font_name

          if "Pain Point" in text:
            for index in range(len(event['pain_points'])):
              if text == f"Pain Point {index + 1}":
                json_data = json.loads(event['pain_points'][index].replace('=>', ':'))
                data = json_data['text']
                paragraph.text = paragraph.text.replace(text, data if data else 'Pain Point: N/A')
                paragraph.text = paragraph.text.replace('%', '')
                if color:
                  paragraph.runs[0].font.color.rgb = color
                paragraph.runs[0].font.size = size
                paragraph.runs[0].font.bold = bold
                paragraph.runs[0].font.name = font_name

            for index in range(len(event['pain_points'])):
              if text == f"Pain Point {index + 1} Data Point 1":
                json_data = json.loads(event['pain_points'][index].replace('=>', ':'))
                data = json_data['illustration']
                paragraph.text = paragraph.text.replace(text, data if data else 'Pain Point Data Point: N/A')
                paragraph.text = paragraph.text.replace('%', '')
                if color:
                  paragraph.runs[0].font.color.rgb = color
                paragraph.runs[0].font.size = size
                paragraph.runs[0].font.bold = bold
                paragraph.runs[0].font.name = font_name

          if "Transition" in text:
            for index in range(len(event['transitions'])):
              if index < len(event['transitions']) and text == f"Transition to {trans_num[index]} Insight":
                paragraph.text = paragraph.text.replace(text, event['transitions'][index] if event['transitions'][index] else 'Transition: N/A')
                paragraph.text = paragraph.text.replace('%', '')
                if color:
                  paragraph.runs[0].font.color.rgb = color
                paragraph.runs[0].font.size = size
                paragraph.runs[0].font.bold = bold
                paragraph.runs[0].font.name = font_name
              elif (index == len(event['transitions']) - 1 and event['closing_transition']) or text == 'Transition To Closing':
                paragraph.text = paragraph.text.replace(text, event['closing_transition'] if event['closing_transition'] else 'Transition To Closing: N/A')
                paragraph.text = paragraph.text.replace('%', '')
                if color:
                  paragraph.runs[0].font.color.rgb = color
                paragraph.runs[0].font.size = size
                paragraph.runs[0].font.bold = bold
                paragraph.runs[0].font.name = font_name

          if text == 'Challenge Question':
            paragraph.text = paragraph.text.replace(text, event['challenge_question'] if event['challenge_question'] else 'Challenge Question: N/A')
            paragraph.text = paragraph.text.replace('%', '')
            if color:
              paragraph.runs[0].font.color.rgb = color
            paragraph.runs[0].font.size = size
            paragraph.runs[0].font.bold = bold
            paragraph.runs[0].font.name = font_name

          elif text == 'Problem':
            paragraph.text = paragraph.text.replace(text, event['problem'] if event['problem'] else 'Problem: N/A')
            paragraph.text = paragraph.text.replace('%', '')
            if color:
              paragraph.runs[0].font.color.rgb = color
            paragraph.runs[0].font.size = size
            paragraph.runs[0].font.bold = bold
            paragraph.runs[0].font.name = font_name

          elif text == 'Summarize Problem':
            paragraph.text = paragraph.text.replace(text, event['summarize_problem'] if event['summarize_problem'] else 'Summarize Problem: N/A')
            paragraph.text = paragraph.text.replace('%', '')
            if color:
              paragraph.runs[0].font.color.rgb = color
            paragraph.runs[0].font.size = size
            paragraph.runs[0].font.bold = bold
            paragraph.runs[0].font.name = font_name

          elif text == 'Call to Action':
            paragraph.text = paragraph.text.replace(text, event['call_to_action'] if event['call_to_action'] else 'Call to Action: N/A')
            paragraph.text = paragraph.text.replace('%', '')
            if color:
              paragraph.runs[0].font.color.rgb = color
            paragraph.runs[0].font.size = size
            paragraph.runs[0].font.bold = bold
            paragraph.runs[0].font.name = font_name

          elif text == 'Closing Statement':
            paragraph.text = paragraph.text.replace(text, event['closing_statement'] if event['closing_statement'] else 'Closing Statement: N/A')
            paragraph.text = paragraph.text.replace('%', '')
            if color:
              paragraph.runs[0].font.color.rgb = color
            paragraph.runs[0].font.size = size
            paragraph.runs[0].font.bold = bold
            paragraph.runs[0].font.name = font_name

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
