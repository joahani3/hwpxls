import React from "react";
import ReactDOM from "react-dom/client";
import App from "./App";
import "./index.css";
import { attachConsole } from "@tauri-apps/plugin-log";

attachConsole(); // 프론트 console.* → 로그 파일에도 기록

ReactDOM.createRoot(document.getElementById("root")!).render(
  <React.StrictMode>
    <App />
  </React.StrictMode>
);
