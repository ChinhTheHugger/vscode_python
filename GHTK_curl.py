import requests
import json

headers = {
    'apikey': 'XXXNMvHSw5nrqXFjNeXPkZZZ',
    'Content-Type': 'application/json',
}

json_data = {
    'name': '75 Âu Cơ',
}

response = requests.post('https://model-deployment.ghtklab.com/api/parse-address-v2', headers=headers, json=json_data)

# print(response.json()) # print response in json format

# print(json.dumps(response.json(), sort_keys=True, ensure_ascii=False, indent=4, separators=(',', ': '))) # pretty print response in json format

# parsed_data = response.json()
# data = parsed_data['data']
# print(json.dumps(data, sort_keys=True, ensure_ascii=False, indent=4, separators=(',', ': '))) # pretty print nested json 'data' in response

parsed_data = response.json()
datas = parsed_data['data']
for data in datas:
    print(json.dumps(data, sort_keys=True, ensure_ascii=False, indent=4, separators=(',', ': '))) # pretty print x  each nested json in 'data'



#-----------------------------------------------------------------------------------------------------------------------------------------------------

# Note: json_data will not be serialized by requests
# exactly as it was in the original request.
#data = '{\n"name": "số 8 Phạm Hùng, Mễ Trì, Nam Từ Liêm, Hà Nội"\n}'.encode()
#response = requests.post('https://model-deployment.ghtklab.com/api/parse-address-v2', headers=headers, data=data)

# curl --location 'https://model-deployment.ghtklab.com/api/parse-address-v2' \
# --header 'apikey: XXXNMvHSw5nrqXFjNeXPkZZZ' \
# --header 'Content-Type: application/json' \
# --data '{
# "name": "số 8 Phạm Hùng, Mễ Trì, Nam Từ Liêm, Hà Nội"
# }'