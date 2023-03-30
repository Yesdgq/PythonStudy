#!/usr/bin/python3

import requests

# GET
r = requests.get('https://api.github.com/user', auth=('user', 'pass'))
print(r.status_code)


print(b'Python is a programming language' in r.content)
print(r.headers['content-type'])
print(r.encoding)
print(r.json())


# POST