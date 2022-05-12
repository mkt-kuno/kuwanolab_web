FROM python:latest

RUN mkdir -p /opt/app

COPY requirements.txt /opt/app/requirements.txt

RUN pip install --upgrade pip && pip install -r /opt/app/requirements.txt && rm /opt/app/requirements.txt

WORKDIR /opt/app

EXPOSE 80

#CMD ["gunicorn", "main:app", "--reload", "-w", "4", "-k", "uvicorn.workers.UvicornWorker", "-b", "0.0.0.0:80"]
#CMD ["uvicorn", "main:app", "--host", "0.0.0.0", "--port", "80"]
CMD ["uvicorn", "main:app", "--reload", "--host", "0.0.0.0", "--port", "80"]

