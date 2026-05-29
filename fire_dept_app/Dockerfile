FROM python:3.10-slim

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY app.py .
COPY static ./static

# Default data files are baked into /app/seed; the entrypoint copies any
# that are missing into the /app/data volume on first run, so the named
# volume persists scraped data across redeploys without starting empty.
COPY data ./seed
COPY entrypoint.sh .
RUN chmod +x entrypoint.sh

EXPOSE 5555

ENTRYPOINT ["./entrypoint.sh"]
