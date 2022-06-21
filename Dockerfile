FROM python:3.10
RUN mkdir /app
ADD . /app
WORKDIR /app
RUN pip3 install -r requirements.txt
CMD ["python", "vtu_result.py"]