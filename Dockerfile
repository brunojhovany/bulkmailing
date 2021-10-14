FROM python:3.8 as builder

WORKDIR /usr/src/app

COPY . .

RUN pip install -r requirements.txt

CMD ["pthon","app.py"]
