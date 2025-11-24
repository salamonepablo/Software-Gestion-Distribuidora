openssl genrsa -out 2Quilplac2021.key 2048
openssl req -new -key 2Quilplac2021.key -subj "/C=AR/O=Quilplac/CN=2Certificado2021/serialNumber=CUIT 30708432543" -out 2Quilplac2021.csr
pause