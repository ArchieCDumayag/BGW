package com.jbgw.technician;

import android.app.Activity;
import android.os.Bundle;
import android.text.InputType;
import android.widget.ArrayAdapter;
import android.widget.Button;
import android.widget.EditText;
import android.widget.LinearLayout;
import android.widget.ScrollView;
import android.widget.Spinner;
import android.widget.TextView;

import org.json.JSONArray;
import org.json.JSONObject;

public class MainActivity extends Activity {
    private LinearLayout content;
    private TextView status;
    private int technicianId = 1;
    private String role = "";
    private String displayName = "";

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        showLogin();
    }

    private void showLogin() {
        LinearLayout root = new LinearLayout(this);
        root.setOrientation(LinearLayout.VERTICAL);
        root.setPadding(32, 48, 32, 32);

        TextView title = row("3JBGW Billing Login");
        title.setTextSize(24);
        root.addView(title);

        Spinner roleSpinner = new Spinner(this);
        roleSpinner.setAdapter(new ArrayAdapter<>(this, android.R.layout.simple_spinner_dropdown_item,
                new String[]{"Technician", "Collector", "Admin"}));
        root.addView(roleSpinner);

        EditText username = input("Username", InputType.TYPE_CLASS_TEXT);
        root.addView(username);

        EditText password = input("Password", InputType.TYPE_CLASS_TEXT | InputType.TYPE_TEXT_VARIATION_PASSWORD);
        root.addView(password);

        status = row("");
        root.addView(status);

        Button login = new Button(this);
        login.setText("Login");
        login.setOnClickListener(v -> login(
                username.getText().toString(),
                password.getText().toString(),
                roleSpinner.getSelectedItem().toString()));
        root.addView(login);

        setContentView(root);
    }

    private void login(String username, String password, String selectedRole) {
        runAsync(() -> {
            JSONObject request = new JSONObject();
            request.put("username", username);
            request.put("password", password);
            request.put("role", selectedRole);

            JSONObject response = new JSONObject(ApiClient.postJson("/api/auth/login", request.toString()));
            role = response.optString("role");
            displayName = response.optString("displayName");
            technicianId = response.optInt("technicianId", 1);

            runOnUiThread(() -> {
                if ("Admin".equalsIgnoreCase(role)) {
                    showAdminNotice();
                } else {
                    showTechnicianHome();
                    loadClients();
                }
            });
        });
    }

    private void showAdminNotice() {
        LinearLayout root = baseShell("Admin Login");
        content.addView(row("Admin account verified.\nUse the web dashboard for admin tools:\n" + ApiClient.BASE_URL + "/Admin/Dashboard"));
        Button logout = new Button(this);
        logout.setText("Logout");
        logout.setOnClickListener(v -> showLogin());
        content.addView(logout);
        setContentView(root);
    }

    private void showTechnicianHome() {
        LinearLayout root = baseShell(displayName + " - " + role);

        LinearLayout actions = new LinearLayout(this);
        actions.setOrientation(LinearLayout.HORIZONTAL);

        Button clientsButton = new Button(this);
        clientsButton.setText("Clients");
        clientsButton.setOnClickListener(v -> loadClients());
        actions.addView(clientsButton);

        Button jobsButton = new Button(this);
        jobsButton.setText("Jobs");
        jobsButton.setOnClickListener(v -> loadJobs());
        actions.addView(jobsButton);

        Button logoutButton = new Button(this);
        logoutButton.setText("Logout");
        logoutButton.setOnClickListener(v -> showLogin());
        actions.addView(logoutButton);

        root.addView(actions, 1);
        setContentView(root);
    }

    private LinearLayout baseShell(String heading) {
        LinearLayout root = new LinearLayout(this);
        root.setOrientation(LinearLayout.VERTICAL);
        root.setPadding(24, 24, 24, 24);

        status = new TextView(this);
        status.setText(heading);
        status.setTextSize(20);
        root.addView(status);

        ScrollView scroll = new ScrollView(this);
        content = new LinearLayout(this);
        content.setOrientation(LinearLayout.VERTICAL);
        scroll.addView(content);
        root.addView(scroll, new LinearLayout.LayoutParams(
                LinearLayout.LayoutParams.MATCH_PARENT,
                0,
                1));
        return root;
    }

    private void loadClients() {
        runAsync(() -> {
            String response = ApiClient.get("/api/technician/clients?technicianId=" + technicianId);
            JSONArray clients = new JSONArray(response);
            runOnUiThread(() -> showClients(clients));
        });
    }

    private void loadJobs() {
        runAsync(() -> {
            String response = ApiClient.get("/api/technician/jobs?technicianId=" + technicianId);
            JSONArray jobs = new JSONArray(response);
            runOnUiThread(() -> showJobs(jobs));
        });
    }

    private void showClients(JSONArray clients) {
        content.removeAllViews();
        status.setText("Assigned clients: " + clients.length());
        for (int i = 0; i < clients.length(); i++) {
            JSONObject client = clients.optJSONObject(i);
            if (client == null) {
                continue;
            }

            LinearLayout item = new LinearLayout(this);
            item.setOrientation(LinearLayout.VERTICAL);
            item.setPadding(0, 18, 0, 18);
            item.addView(row(client.optString("name") + "\n" +
                    client.optString("area") + " " + client.optString("zone") + "\n" +
                    "Contact: " + client.optString("contact") + "\n" +
                    "PPPoE: " + client.optString("pppoeUsername")));

            Button details = new Button(this);
            details.setText("Location / PPPoE Status");
            int clientId = client.optInt("id");
            details.setOnClickListener(v -> loadClientDetails(clientId));
            item.addView(details);
            content.addView(item);
        }
    }

    private void loadClientDetails(int clientId) {
        runAsync(() -> {
            JSONObject response = new JSONObject(ApiClient.get("/api/technician/clients/" + clientId));
            JSONObject client = response.optJSONObject("client");
            JSONObject billingRule = response.optJSONObject("billingRule");
            JSONObject pppoe = response.optJSONObject("pppoe");
            JSONArray payments = response.optJSONArray("payments");
            runOnUiThread(() -> {
                content.removeAllViews();
                status.setText("Client details");
                StringBuilder details = new StringBuilder();
                details.append(client.optString("name")).append("\n")
                        .append("Location: ").append(client.optString("address")).append("\n")
                        .append("Contact: ").append(client.optString("contact")).append("\n")
                        .append("PPPoE: ").append(pppoe.optString("username")).append("\n")
                        .append("PPPoE Status: ").append(pppoe.optString("status")).append("\n")
                        .append("Type: ").append(client.optString("billingType")).append("\n")
                        .append("Current Bill: PHP ").append(client.optString("bills")).append("\n")
                        .append("Balance: PHP ").append(client.optString("balance")).append("\n")
                        .append("Advance: PHP ").append(client.optString("advance"));

                if (billingRule != null) {
                    details.append("\n")
                            .append("Due: ")
                            .append(billingRule.optString("scheduleLabel"))
                            .append(" (")
                            .append(billingRule.optString("nextDueDate"))
                            .append(")");
                    if (billingRule.optBoolean("hasEarlyDiscount")) {
                        details.append("\nXentronet discount: PHP ")
                                .append(billingRule.optString("earlyDiscountAmount"))
                                .append(" before ")
                                .append(billingRule.optString("discountDeadline"))
                                .append("\nDiscounted Bill: PHP ")
                                .append(billingRule.optString("discountedCurrentBill"));
                    }
                }

                if (payments != null && payments.length() > 0) {
                    details.append("\n\nRecent payments:");
                    int paymentCount = Math.min(payments.length(), 5);
                    for (int i = 0; i < paymentCount; i++) {
                        JSONObject payment = payments.optJSONObject(i);
                        if (payment == null) {
                            continue;
                        }

                        details.append("\n")
                                .append(payment.optString("paidOn"))
                                .append(" - PHP ")
                                .append(payment.optString("amount"))
                                .append(" (")
                                .append(payment.optString("method"))
                                .append(")");
                    }
                } else {
                    details.append("\n\nRecent payments: none");
                }

                content.addView(row(details.toString()));
                Button back = new Button(this);
                back.setText("Back to Clients");
                back.setOnClickListener(v -> loadClients());
                content.addView(back);
            });
        });
    }

    private void showJobs(JSONArray jobs) {
        content.removeAllViews();
        status.setText("Jobs: " + jobs.length());
        for (int i = 0; i < jobs.length(); i++) {
            JSONObject wrapper = jobs.optJSONObject(i);
            JSONObject job = wrapper == null ? null : wrapper.optJSONObject("job");
            JSONObject client = wrapper == null ? null : wrapper.optJSONObject("client");
            if (job == null) {
                continue;
            }
            LinearLayout item = new LinearLayout(this);
            item.setOrientation(LinearLayout.VERTICAL);
            item.setPadding(0, 18, 0, 18);

            item.addView(row(job.optString("type") + " - " + job.optString("status") + "\n" +
                    (client == null ? "Client ID " + job.optInt("clientId") : client.optString("name")) + "\n" +
                    job.optString("remarks")));

            EditText remarks = input("Job remarks", InputType.TYPE_CLASS_TEXT);
            item.addView(remarks);

            int jobId = job.optInt("id");

            Button addRemarks = new Button(this);
            addRemarks.setText("Add Remarks");
            addRemarks.setOnClickListener(v -> addRemarks(jobId, remarks.getText().toString()));
            item.addView(addRemarks);

            Button done = new Button(this);
            done.setText("Mark Done");
            done.setOnClickListener(v -> completeJob(jobId, remarks.getText().toString()));
            item.addView(done);
            content.addView(item);
        }
    }

    private void addRemarks(int jobId, String remarks) {
        runAsync(() -> {
            JSONObject request = new JSONObject();
            request.put("remarks", remarks);
            ApiClient.postJson("/api/technician/jobs/" + jobId + "/remarks", request.toString());
            loadJobs();
        });
    }

    private void completeJob(int jobId, String remarks) {
        runAsync(() -> {
            JSONObject request = new JSONObject();
            request.put("remarks", remarks.isEmpty() ? "Completed from Android app" : remarks);
            ApiClient.postJson("/api/technician/jobs/" + jobId + "/done", request.toString());
            loadJobs();
        });
    }

    private EditText input(String hint, int type) {
        EditText input = new EditText(this);
        input.setHint(hint);
        input.setInputType(type);
        return input;
    }

    private TextView row(String text) {
        TextView view = new TextView(this);
        view.setText(text);
        view.setTextSize(16);
        view.setPadding(0, 12, 0, 12);
        return view;
    }

    private void runAsync(Task task) {
        if (status != null) {
            status.setText("Loading...");
        }
        new Thread(() -> {
            try {
                task.run();
            } catch (Exception ex) {
                runOnUiThread(() -> {
                    if (status != null) {
                        status.setText("Error: " + ex.getMessage());
                    }
                });
            }
        }).start();
    }

    private interface Task {
        void run() throws Exception;
    }
}
