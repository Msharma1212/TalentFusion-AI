import { execSync } from "child_process";

try {
  console.log("Calling local API health endpoint...");
  const res = execSync("curl -s http://localhost:3000/api/health").toString().trim();
  console.log("Health Check Result:", res);
} catch (err: any) {
  console.log("Error contacting local server on port 3000:", err.message);
}
