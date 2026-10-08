const express = require('express');
const cors = require('cors');
const { Pool } = require('pg');
require('dotenv').config();

const app = express();
app.use(cors());
app.use(express.json());

const pool = new Pool({
  connectionString: process.env.DATABASE_URL,
  ssl: {
    rejectUnauthorized: false
  }
});

// Initialize database table
const initDB = async () => {
  try {
    await pool.query(`
      CREATE TABLE IF NOT EXISTS offchain_orders (
        id VARCHAR(255) PRIMARY KEY,
        image_url TEXT,
        customer_name VARCHAR(255),
        customer_email VARCHAR(255),
        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
      );
    `);
    console.log("Database initialized");
  } catch (err) {
    console.error("Error initializing DB:", err);
  }
};
initDB();

// API to get order off-chain details
app.get('/api/orders/:id', async (req, res) => {
  try {
    const { rows } = await pool.query('SELECT * FROM offchain_orders WHERE id = $1', [req.params.id]);
    if (rows.length > 0) {
      res.json(rows[0]);
    } else {
      res.status(404).json({ error: 'Order not found in DB' });
    }
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// API to save order off-chain details
app.post('/api/orders', async (req, res) => {
  const { id, image_url, customer_name, customer_email } = req.body;
  try {
    await pool.query(
      `INSERT INTO offchain_orders (id, image_url, customer_name, customer_email) 
       VALUES ($1, $2, $3, $4) 
       ON CONFLICT (id) DO UPDATE 
       SET image_url = $2, customer_name = $3, customer_email = $4`,
      [id, image_url, customer_name, customer_email]
    );
    res.json({ success: true, message: 'Saved to Neon DB' });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

const PORT = process.env.PORT || 5000;
app.listen(PORT, () => {
  console.log(`Server running on port ${PORT}`);
});
