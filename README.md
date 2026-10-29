# CovantisTrust: Decentralized Cold-Chain Supply Chain Protocol

**BCSE329L - Blockchain and Distributed Ledger Technology Course Project**  
**School of Computer Science and Engineering (SCOPE), Vellore Institute of Technology (VIT)**

* **Live Interactive Web3 Staging Deployment:** [https://frontend-nine-henna-77.vercel.app](https://frontend-nine-henna-77.vercel.app)
* **Public GitHub Repository:** [https://github.com/arindam199/CovantisTrust](https://github.com/arindam199/CovantisTrust)

---

## 1. Project Title
**CovantisTrust: Industrial-Grade Decentralized Cold-Chain Order Tracking and Telemetry Protocol Using Hybrid EVM Smart Contracts and Serverless Relational State**

---

## 2. Technologies Used
* **Blockchain Core:** Solidity `v0.8.19`, Hardhat Development Environment `v2.29.1`, EVM net gas metering.
* **Web3 Integration:** Ethers.js `v6.17.0`, MetaMask Browser Extension (EIP-1193 Standard).
* **Frontend UI / Presentation:** React `19.3.0`, Vite `8.2.2`, Tailwind CSS `v4.3.3`, Lucide React icons.
* **Off-Chain Middleware API:** Node.js `v26.7.0`, Express.js `v5.2.1`, `cors`, `dotenv`.
* **Relational Persistence Layer:** Neon Serverless PostgreSQL (`pg` connection pool `v8.23.1`, Postgres 16).
* **Automated Testing & Gas Inspection:** Mocha, Chai, Hardhat Gas Reporter.
* **Process Orchestration:** `concurrently`, `kill-port`.

---

## 3. Blockchain / Networks Used
* **Local Development Network:** Hardhat EVM Node (`http://127.0.0.1:8545`)
  * **Chain ID:** `31337`
  * **Block Gas Limit:** `30,000,000`
  * **Mining Mode:** Instantaneous synthetic block synthesis (0-latency local testing).
* **Public Ethereum Testnet:** Ethereum Sepolia Testnet
  * **Chain ID:** `11155111`
  * **Consensus:** Proof-of-Stake (PoS), EIP-1559 Dynamic Fee Market.
* **Layer-2 Rollup (Simulation Target):** Arbitrum Nitro L2 Rollup (`Chain ID: 421614`).

---

## 4. Dependencies Required
Before running the project, ensure your workstation has:
1. **Node.js:** `v20.0.0` or higher (`node -v`).
2. **npm:** `v10.0.0` or higher (`npm -v`).
3. **MetaMask Extension:** Installed on your web browser.
4. **PostgreSQL Database:** Connection string to a Neon Serverless PostgreSQL instance (provided in `backend/.env`).

---

## 5. Setup and Installation Procedure

Clone the repository and install dependencies across the monorepo:

```bash
# 1. Clone repository
git clone https://github.com/arindam199/CovantisTrust.git
cd CovantisTrust

# 2. Install root orchestrator dependencies
npm install

# 3. Install frontend dependencies
cd frontend
npm install
cd ..

# 4. Install backend middleware dependencies
cd backend
npm install
cd ..
```

---

## 6. Steps to Run the Project

The workspace includes a single-command master orchestrator script (`npm run start:all`) that launches all services concurrently:

```bash
npm run start:all
```

This single command automatically:
1. Executes `kill-port 8545` to prevent dangling socket collisions.
2. Boots the Hardhat EVM local node at `http://127.0.0.1:8545`.
3. Compiles and deploys `OrderTracking.sol` via `scripts/deploy.js`.
4. Dynamically injects the landed contract address into `frontend/src/App.jsx`.
5. Starts the Express/PostgreSQL API server on port `5000`.
6. Launches the Vite development server with hot-module reloading at `http://localhost:5173`.

---

## 7. Smart Contract Deployment Details

* **Smart Contract Source File:** `contracts/OrderTracking.sol`
* **Solidity Compiler Target:** `0.8.19` (Optimizer enabled, 200 runs)
* **Deployment Script:** `scripts/deploy.js`
* **Default Local Contract Address:** `0x5FbDB2315678afecb367f032d93F642f64180aa3` (Hardhat Account #0 default deployment)
* **Application Binary Interface (ABI):** Generated automatically in `frontend/src/artifacts/contracts/OrderTracking.sol/OrderTracking.json`.

---

## 8. Main Smart Contract Functions

| Function Name | State Mutability | Access Control Policy | Description |
|---|---|---|---|
| `createOrder(string _productDetails, address _buyer)` | State Mutating | Public (Callable by seller) | Instantiates a new consignment record with status `Created (0)`, genesis timestamp, and emits `OrderCreated`. |
| `updateOrderStatus(uint256 _id, OrderStatus _status)` | State Mutating | `require(msg.sender == orders[_id].seller)` | Advances consignment through FSM states (`Processing`, `Shipped`, `InTransit`, `Delivered`). |
| `logIoTReading(uint256 _id, int256 _temperature, string _location)` | State Mutating | Public (IoT Sensor Telematics) | Logs temperature telemetry; autonomously trips `conditionViolated = true` if temperature breaches GDP safe band (`2°C to 8°C`). |
| `cancelOrder(uint256 _id)` | State Mutating | Seller OR Buyer (`Created` / `Processing` only) | Revokes active order before shipment handoff; sets state to `Canceled (5)`. |
| `getOrder(uint256 _id)` | View (Read-Only) | Public | Returns complete `Order` struct tuple (0 gas execution via `eth_call`). |
| `getOrderReadings(uint256 _id)` | View (Read-Only) | Public | Returns full historical array of timestamped `IoTReading` structs for audit. |

---

## 9. Automated Testing Suite

To execute the Mocha/Chai automated verification test suite:

```bash
npx hardhat test
```

Verification covers:
* `TC-01`: Valid consignment instantiation and counter increment.
* `TC-02`: Authorized status updates by seller.
* `TC-03`: Unauthorized status mutation reversion (`Only seller can update status`).
* `TC-04`: Compliant IoT temperature logging (`conditionViolated = false`).
* `TC-05`: Autonomous GDP thermal violation trigger (`conditionViolated = true`).
* `TC-06`: Permissible order cancellation during early stages.
* `TC-07`: Illegal in-transit cancellation reversion (`Cannot cancel now`).

---

## 10. Authors & Course Information
* **Lead Researcher / Student:** Arindam (Reg. No: 22BCE10482)
* **Department:** School of Computer Science and Engineering (SCOPE)
* **Institution:** Vellore Institute of Technology (VIT), Vellore, India
* **Course:** BCSE329L - Blockchain and Distributed Ledger Technology
