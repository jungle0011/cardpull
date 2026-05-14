FROM node:18-slim

WORKDIR /app

ENV NODE_ENV=production
ENV PORT=3000
ENV PLAYWRIGHT_BROWSERS_PATH=/ms-playwright

COPY package*.json ./
RUN npm ci
RUN npx playwright install chromium --with-deps

COPY . .
RUN mkdir -p /tmp/uploads /tmp/output

EXPOSE 3000

CMD ["npm", "start"]
