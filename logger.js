// ไฟล์นี้จะช่วยดักจับ Error 400 จากทุกที่ในโปรเจกต์
process.on('unhandledRejection', (reason) => {
    console.log("---------- DETAILED ERROR ----------");
    if (reason && reason.response) {
        console.error("Status:", reason.response.status);
        console.error("Data from API:", JSON.stringify(reason.response.data, null, 2));
    } else {
        console.error("Error:", reason);
    }
    console.log("------------------------------------");
});
console.log("✅ Debug Logger is active.");
