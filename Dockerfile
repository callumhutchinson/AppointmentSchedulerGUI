# syntax=docker/dockerfile:1

FROM python:2.7-slim-buster

WORKDIR /app

COPY requirements.txt requirements.txt
RUN pip install -r requirements.txt

COPY . .

CMD [ "python", "appointment_scheduler.py"]
