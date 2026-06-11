from mailjet_rest import Client
mailjet = Client(auth=('54e7bf3cfcc10de5367921736732a31d', 'd16abd9ce50466e1e0fb4bce1720b4b9'), version='v3')

# Tenta apagar o ID que ficou "preso"
result = mailjet.campaigndraft.delete(id=000000000)
print(f"Resultado: {result.status_code}") # 204 significa que apagou mesmo