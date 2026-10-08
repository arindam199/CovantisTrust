import hre from "hardhat";
import fs from "fs";

async function main() {
  const OrderTracking = await hre.ethers.getContractFactory("OrderTracking");
  const orderTracking = await OrderTracking.deploy();

  await orderTracking.waitForDeployment();

  const contractAddress = await orderTracking.getAddress();
  console.log("OrderTracking deployed to:", contractAddress);

  // Read App.jsx and replace the contract address
  const appPath = "./frontend/src/App.jsx";
  let appCode = fs.readFileSync(appPath, "utf-8");
  
  // Find the line with contractAddress and replace it
  appCode = appCode.replace(
    /const contractAddress = "0x[a-fA-F0-9]{40}";/,
    `const contractAddress = "${contractAddress}";`
  );
  
  fs.writeFileSync(appPath, appCode);
  console.log("Updated contractAddress in App.jsx");
}

main().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
