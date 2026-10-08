import { expect } from "chai";
import hre from "hardhat";

describe("OrderTracking", function () {
  let OrderTracking;
  let orderTracking;
  let owner;
  let buyer;
  let otherAccount;

  beforeEach(async function () {
    [owner, buyer, otherAccount] = await hre.ethers.getSigners();
    OrderTracking = await hre.ethers.getContractFactory("OrderTracking");
    orderTracking = await OrderTracking.deploy();
  });

  describe("Order Creation", function () {
    it("Should create a new order", async function () {
      const tx = await orderTracking.createOrder("Laptop MacBook Pro", buyer.address);
      const receipt = await tx.wait();

      const orderCount = await orderTracking.orderCount();
      expect(orderCount).to.equal(1);

      const order = await orderTracking.getOrder(1);
      expect(order.productDetails).to.equal("Laptop MacBook Pro");
      expect(order.seller).to.equal(owner.address);
      expect(order.buyer).to.equal(buyer.address);
      expect(order.status).to.equal(0n); // Created
    });
  });

  describe("Order Status Update", function () {
    beforeEach(async function () {
      await orderTracking.createOrder("Laptop MacBook Pro", buyer.address);
    });

    it("Should allow seller to update order status", async function () {
      await orderTracking.updateOrderStatus(1, 1); // Processing
      let order = await orderTracking.getOrder(1);
      expect(order.status).to.equal(1n);

      await orderTracking.updateOrderStatus(1, 2); // Shipped
      order = await orderTracking.getOrder(1);
      expect(order.status).to.equal(2n);
    });

    it("Should not allow non-seller to update order status", async function () {
      await expect(
        orderTracking.connect(buyer).updateOrderStatus(1, 1)
      ).to.be.revertedWith("Only seller can perform this action");
    });
  });

  describe("Order Cancellation", function () {
    beforeEach(async function () {
      await orderTracking.createOrder("Laptop MacBook Pro", buyer.address);
    });

    it("Should allow buyer to cancel order if it is in Created state", async function () {
      await orderTracking.connect(buyer).cancelOrder(1);
      const order = await orderTracking.getOrder(1);
      expect(order.status).to.equal(5n); // Canceled
    });

    it("Should allow seller to cancel order if it is in Created state", async function () {
      await orderTracking.cancelOrder(1);
      const order = await orderTracking.getOrder(1);
      expect(order.status).to.equal(5n); // Canceled
    });

    it("Should not allow cancellation if shipped", async function () {
      await orderTracking.updateOrderStatus(1, 2); // Shipped
      await expect(
        orderTracking.connect(buyer).cancelOrder(1)
      ).to.be.revertedWith("Cannot cancel at this stage");
    });
  });
});
