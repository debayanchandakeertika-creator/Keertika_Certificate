# import requests
# import json

# url = "http://127.0.0.1:8000/storepostorderjson"
# headers = {'Content-Type': 'application/json'}

# # Example JSON payload
# payload = {'name': 'John Doe', 'essay': ""}

# # Convert the payload to JSON
# json_payload = json.dumps(payload)

# # Make the POST request with the JSON payload
# response = requests.post(url, headers=headers, data=json_payload)

# # Print the response from the server
# print(response.text)

import json
import requests
import io
# Define the endpoint URL and JSON payload
url = 'http://127.0.0.1:8000/storepostorderjson'
# payload = {'name': 'John Doe', 'email': 'johndoe@example.com'}



json_data = {
    "name": "John",
    "age": 30,
    "email": "john@example.com"
}

# Encode the JSON data as a string in multipart format
json_data_encoded = json.dumps(json_data).encode('utf-8')
files = {'json': ('json_data', json_data_encoded, 'application/json')}

# Send the API request with the JSON data in multipart format
response = requests.post(url, files=files)

# Check the API response
if response.status_code == 200:
    print("API request successful")
else:
    print("API request failed with error code:", response.status_code)


# # Create a dictionary to store the JSON data
# json_data = {'key': 'value'}

# # Convert the JSON data to a string
# json_str = json.dumps(json_data)

# # Create a binary file-like object with the JSON data
# json_file = io.BytesIO(json_str.encode('utf-8'))

# # Create a dictionary to store the multipart data
# multipart_data = {
#     'json': ('data.json', json_file, 'application/json')
# }

# # Make the API request with the multipart data
# response = requests.post(url, files=multipart_data)

# # Print the response status code and JSON response data
# print(response.status_code)
# print(response.json())

# import requests
# import json

# url = 'https://example.com/api/endpoint'
# headers = {'Content-Type': 'multipart/form-data  boundary=----WebKitFormBoundary7MA4YWxkTrZu0gW', }
# data = {'json': ('data.json', json.dumps({'key': 'value'}), 'application/json')}

# response = requests.post(url, headers=headers, files=data)

# print(response.status_code)
# print(response.json())


# # Create a new multipart request
# request = requests.Request('POST', url)
# json_payload = json.dumps(payload)
# request_data = {'payload': json_payload}
# multipart_data = requests.urllib3.filepost.encode_multipart_formdata(request_data)
# print(multipart_data)
# # Set the Content-Type header to multipart/form-data
# headers = {'Content-Type': multipart_data.content_type}

# # Send the HTTP POST request
# response = requests.post(url, headers=headers, data=multipart_data)

# # Check the response status code
# if response.status_code == 200:
#     # Success: parse the response body as JSON
#     response_data = json.loads(response.text)
#     print(response_data)
# else:
#     # Error: print the response status code and message
#     print(f'Error: {response.status_code} - {response.reason}')




# import 'dart:convert';
# import 'package:http/http.dart' as http;

# void main() async {
#   // Define the endpoint URL and JSON payload
#   final url = Uri.parse('https://example.com/api');
#   final payload = {'name': 'John Doe', 'email': 'johndoe@example.com'};

#   // Encode the payload as a JSON string
#   final jsonPayload = jsonEncode(payload);

#   // Send the HTTP POST request with the JSON payload in the request body
#   final response = await http.post(url, body: jsonPayload, headers: {'Content-Type': 'application/json'});

#   // Check the response status code
#   if (response.statusCode == 200) {
#     // Success: parse the response body as JSON
#     final responseBody = jsonDecode(response.body);
#     print(responseBody);
#   } else {
#     // Error: print the response status code and message
#     print('Error: ${response.statusCode} - ${response.reasonPhrase}');
#   }
# }


# import 'dart:convert';
# import 'dart:io';
# import 'package:http/http.dart' as http;

# void main() async {
#   // Define the endpoint URL and JSON payload
#   final url = Uri.parse('https://example.com/api');
#   final payload = {'name': 'John Doe', 'email': 'johndoe@example.com'};

#   // Create a new multipart request
#   final request = http.MultipartRequest('POST', url);

#   // Add the JSON payload as a new part
#   request.fields['payload'] = jsonEncode(payload);

#   // Send the HTTP POST request
#   final response = await request.send();

#   // Check the response status code
#   if (response.statusCode == 200) {
#     // Success: parse the response body as JSON
#     final responseBody = await response.stream.bytesToString();
#     print(responseBody);
#   } else {
#     // Error: print the response status code and message
#     print('Error: ${response.statusCode} - ${response.reasonPhrase}');
#   }
# }
