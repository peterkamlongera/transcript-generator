import React from "react";
import TranscriptGenerator from "../components/transcript-generator";

const HomePage = () => {
  return (
    <>
      <h1>Boyfriend of The Year Award</h1>
      <div
        style={{
          backgroundColor: "#007BFF",
          minHeight: "100vh",
          display: "flex",
          justifyContent: "center",
          alignItems: "center",
          padding: "20px",
        }}
      >
        <TranscriptGenerator />
      </div>
    </>
  );
};

export default HomePage;
