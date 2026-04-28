package com.jbgw.technician;

import java.io.BufferedReader;
import java.io.InputStreamReader;
import java.io.OutputStream;
import java.net.HttpURLConnection;
import java.net.URL;
import java.nio.charset.StandardCharsets;

public final class ApiClient {
    public static final String BASE_URL = "http://10.0.2.2:5001";

    private ApiClient() {
    }

    public static String get(String path) throws Exception {
        HttpURLConnection connection = open(path, "GET");
        return read(connection);
    }

    public static String postJson(String path, String json) throws Exception {
        HttpURLConnection connection = open(path, "POST");
        connection.setRequestProperty("Content-Type", "application/json");
        connection.setDoOutput(true);
        try (OutputStream output = connection.getOutputStream()) {
            output.write(json.getBytes(StandardCharsets.UTF_8));
        }
        return read(connection);
    }

    private static HttpURLConnection open(String path, String method) throws Exception {
        URL url = new URL(BASE_URL + path);
        HttpURLConnection connection = (HttpURLConnection) url.openConnection();
        connection.setRequestMethod(method);
        connection.setConnectTimeout(8000);
        connection.setReadTimeout(8000);
        return connection;
    }

    private static String read(HttpURLConnection connection) throws Exception {
        int status = connection.getResponseCode();
        try (BufferedReader reader = new BufferedReader(new InputStreamReader(
                status >= 400 ? connection.getErrorStream() : connection.getInputStream(),
                StandardCharsets.UTF_8))) {
            StringBuilder builder = new StringBuilder();
            String line;
            while ((line = reader.readLine()) != null) {
                builder.append(line);
            }
            return builder.toString();
        }
    }
}
