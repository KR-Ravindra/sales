FROM python:3.9-slim

WORKDIR /app

COPY ./requirements.txt /app/requirements.txt

RUN pip3 install -r requirements.txt

COPY . /app

EXPOSE 8105

HEALTHCHECK CMD curl --fail http://localhost:8105/_stcore/health

ENTRYPOINT ["streamlit", "run", "app.py", "--server.port=8105", "--server.address=0.0.0.0"]