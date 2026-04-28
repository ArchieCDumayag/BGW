# 3JBGW Technician Android App

This is a small native Android client scaffold for technicians and collectors.

Point `ApiClient.BASE_URL` in `app/src/main/java/com/jbgw/technician/ApiClient.java` to the billing server, for example:

```java
public static final String BASE_URL = "http://192.168.1.10:5001";
```

For Android emulator, keep:

```java
public static final String BASE_URL = "http://10.0.2.2:5001";
```

Default login accounts:

- Admin: `admin` / `admin123`
- Technician: `tech` / `tech123`
- Collector: `collector` / `collector123`

The app uses these server endpoints:

- `POST /api/auth/login`
- `GET /api/technician/clients?technicianId=1`
- `GET /api/technician/jobs?technicianId=1`
- `GET /api/technician/clients/{id}`
- `GET /api/technician/clients/{id}/pppoe`
- `POST /api/technician/jobs/{id}/remarks`
- `POST /api/technician/jobs/{id}/done`

Technician and collector users can view assigned clients, installation and repair jobs, client contact/location, remarks, completion status, and PPPoE status. Destructive admin actions such as deleting payments or changing system settings are not exposed by the app.
