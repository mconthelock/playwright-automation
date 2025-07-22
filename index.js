const express = require("express");
const app = express();
const port = 3000;
const routes = require("./route");
app.use(express.json());
app.use("/", routes);
app.listen(port, () => {
  console.log(`Automation server listening at http://localhost:${port}`);
});
