// SPDX-License-Identifier: MIT
pragma solidity ^0.8.19;

contract OrderTracking {
    enum OrderStatus { Created, Processing, Shipped, InTransit, Delivered, Canceled }

    struct IoTReading {
        int256 temperature;
        uint256 timestamp;
        string location;
    }

    struct Order {
        uint256 id;
        string productDetails;
        address seller;
        address buyer;
        OrderStatus status;
        uint256 createdAt;
        uint256 updatedAt;
        bool conditionViolated; // for cold chain monitoring
    }

    uint256 public orderCount = 0;
    mapping(uint256 => Order) public orders;
    mapping(uint256 => IoTReading[]) public orderReadings;

    event OrderCreated(uint256 id, string productDetails, address indexed seller, address indexed buyer);
    event OrderStatusUpdated(uint256 id, OrderStatus status);
    event IoTReadingLogged(uint256 id, int256 temperature, string location, bool violated);

    function createOrder(string memory _productDetails, address _buyer) public {
        orderCount++;
        orders[orderCount] = Order(
            orderCount,
            _productDetails,
            msg.sender,
            _buyer,
            OrderStatus.Created,
            block.timestamp,
            block.timestamp,
            false
        );
        emit OrderCreated(orderCount, _productDetails, msg.sender, _buyer);
    }

    function updateOrderStatus(uint256 _id, OrderStatus _status) public {
        require(_id > 0 && _id <= orderCount, "Order does not exist");
        require(msg.sender == orders[_id].seller, "Only seller can update status");
        
        orders[_id].status = _status;
        orders[_id].updatedAt = block.timestamp;
        
        emit OrderStatusUpdated(_id, _status);
    }

    // IoT Sensor Integration for Cold Chain
    function logIoTReading(uint256 _id, int256 _temperature, string memory _location) public {
        require(_id > 0 && _id <= orderCount, "Order does not exist");
        require(orders[_id].status != OrderStatus.Delivered && orders[_id].status != OrderStatus.Canceled, "Order closed");
        
        IoTReading memory reading = IoTReading(_temperature, block.timestamp, _location);
        orderReadings[_id].push(reading);
        
        // Cold Chain threshold: Pharma needs to be between 2°C and 8°C
        bool violated = false;
        if (_temperature > 8 || _temperature < 2) {
            orders[_id].conditionViolated = true;
            violated = true;
        }
        
        emit IoTReadingLogged(_id, _temperature, _location, violated);
    }

    function getOrderReadings(uint256 _id) public view returns (IoTReading[] memory) {
        return orderReadings[_id];
    }

    function getOrder(uint256 _id) public view returns (Order memory) {
        require(_id > 0 && _id <= orderCount, "Order does not exist");
        return orders[_id];
    }

    function cancelOrder(uint256 _id) public {
        require(_id > 0 && _id <= orderCount, "Order does not exist");
        require(msg.sender == orders[_id].seller || msg.sender == orders[_id].buyer, "Only seller or buyer can cancel");
        require(orders[_id].status == OrderStatus.Created || orders[_id].status == OrderStatus.Processing, "Cannot cancel now");

        orders[_id].status = OrderStatus.Canceled;
        orders[_id].updatedAt = block.timestamp;
        
        emit OrderStatusUpdated(_id, OrderStatus.Canceled);
    }
}
