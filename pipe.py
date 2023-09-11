from pipefy import Pipefy
import pandas as pd
import re

token = "eyJ0eXAiOiJKV1QiLCJhbGciOiJIUzUxMiJ9.eyJ1c2VyIjp7ImlkIjozMDE0MDM2MDIsImVtYWlsIjoibWF0ZXVzLmxhZGlrYUBncnVwb21pbmlwcmVjby5jb20uYnIiLCJhcHBsaWNhdGlvbiI6MzAwMjA0MTE5fX0.OyI0ST11L_zGyRQSbyWepOKjZMoRLxzqQClMC8CMVUh4lu9yF_QnVOW4PCNfDgPSQAblBGVhoMsOGf4NUmlPVA"
pipe_id = 303059855
pipefy = Pipefy(token)

cards = pipefy.allCards(pipe_id, filter='{field: "updated_at", operator: gt, value: "2023-08-01T23:50:11-03:00"}')

#cardsIds = re.findall(r'\'id\': \'(\d{9})\'',str(cards))
print(cards)
