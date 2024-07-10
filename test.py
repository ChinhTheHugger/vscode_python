import requests

url = 'https://danhmuchanhchinh.gso.gov.vn/NghiDinh.aspx'

def get_page_source(url: str) -> str:
  result = requests.get(url)
  result.raise_for_status()
  return result.text

print(get_page_source(url))