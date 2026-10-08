import { useState, useEffect } from 'react';
import { ethers } from 'ethers';
import { 
  Package, CheckCircle2, Clock, XCircle, 
  Plus, Wallet, ShieldCheck, ArrowRight, Box, 
  Activity, Hash, MapPin, Copy, ExternalLink,
  ChevronRight, Thermometer, ShieldAlert,
  QrCode, AlertTriangle
} from 'lucide-react';
import OrderTrackingArtifact from './artifacts/contracts/OrderTracking.sol/OrderTracking.json';
import { useToast } from './components/Toast';
// Removed unused ToastProvider import


const contractAddress = "0x5FbDB2315678afecb367f032d93F642f64180aa3";

export default function App() {
  const [account, setAccount] = useState(null);
  const { addToast } = useToast();
  const [contract, setContract] = useState(null);
  const [orders, setOrders] = useState([]);
  
  // Form State
  const [productDetails, setProductDetails] = useState('');
  const [buyerAddress, setBuyerAddress] = useState('');
  const [isCreating, setIsCreating] = useState(false);
  const [activeTab, setActiveTab] = useState('active'); // active, all
  
  // Demo IoT State
  const [simTemp, setSimTemp] = useState('5');
  const [simLocation, setSimLocation] = useState('Port of Loading');

  const statusLabels = ["Registered", "Processing", "In Transit", "Out for Delivery", "Delivered", "Canceled"];
  
  useEffect(() => {
    checkWalletConnection();
  }, []);

  const checkWalletConnection = async () => {
    if (window.ethereum) {
      try {
        const accounts = await window.ethereum.request({ method: 'eth_accounts' });
        if (accounts.length > 0) {
          setAccount(accounts[0]);
          initializeContract(accounts[0]);
        }
      } catch (err) {
        console.error(err);
      }
    }
  };

  const connectWallet = async () => {
    if (window.ethereum) {
      try {
        const accounts = await window.ethereum.request({ method: 'eth_requestAccounts' });
        setAccount(accounts[0]);
        initializeContract(accounts[0]);
      } catch (err) {
        console.error(err);
      }
    } else {
      alert("Please install MetaMask to use this application.");
    }
  };

  const initializeContract = async (userAccount) => {
    const provider = new ethers.BrowserProvider(window.ethereum);
    const signer = await provider.getSigner();
    const orderContract = new ethers.Contract(contractAddress, OrderTrackingArtifact.abi, signer);
    setContract(orderContract);
    fetchOrders(orderContract, userAccount);
  };

  const fetchOrders = async (orderContract, userAccount) => {
    try {
      const count = await orderContract.orderCount();
      const loadedOrders = [];
      for (let i = 1; i <= count; i++) {
        const order = await orderContract.getOrder(i);
        const readings = await orderContract.getOrderReadings(i);
        
        let dbData = {};
        try {
          const res = await fetch(`http://localhost:5000/api/orders/${i}`);
          if (res.ok) {
            dbData = await res.json();
          }
        } catch (e) {
          console.error("DB Fetch Error:", e);
        }

        loadedOrders.push({
          id: order.id.toString(),
          productDetails: order.productDetails,
          seller: order.seller,
          buyer: order.buyer,
          status: Number(order.status),
          conditionViolated: order.conditionViolated,
          readings: readings.map(r => ({
              temp: Number(r.temperature),
              time: new Date(Number(r.timestamp) * 1000).toLocaleTimeString([], {hour: '2-digit', minute:'2-digit'}),
              loc: r.location
          })),
          createdAt: new Date(Number(order.createdAt) * 1000).toLocaleString('en-US', { month: 'short', day: 'numeric', hour: '2-digit', minute: '2-digit' }),
          updatedAt: new Date(Number(order.updatedAt) * 1000).toLocaleString('en-US', { month: 'short', day: 'numeric', hour: '2-digit', minute: '2-digit' }),
          // Off-chain data
          imageUrl: dbData.image_url || 'https://via.placeholder.com/150?text=No+Image',
          customerName: dbData.customer_name || 'N/A',
          customerEmail: dbData.customer_email || 'N/A'
        });
      }
      setOrders(loadedOrders.reverse());
    } catch (err) {
      console.error("Failed to fetch orders:", err);
    }
  };

  const createOrder = async (e) => {
    e.preventDefault();
    if (contract) {
      try {
        setIsCreating(true);
        const tx = await contract.createOrder(productDetails, buyerAddress);
        await tx.wait();
        
        // Wait, the id is not returned directly from createOrder since it's an on-chain event.
        // We will just fetch the latest count and use it. (This is a simplified approach, in production we'd listen to the OrderCreated event).
        const count = await contract.orderCount();
        const orderId = count.toString();
        
        try {
           await fetch('http://localhost:5000/api/orders', {
             method: 'POST',
             headers: { 'Content-Type': 'application/json' },
             body: JSON.stringify({
               id: orderId,
               image_url: 'https://images.unsplash.com/photo-1580674684081-77678518e11a?w=400&q=80',
               customer_name: 'John Doe',
               customer_email: 'john@example.com'
             })
           });
        } catch(err) {
           console.error("DB Save Error:", err);
        }

        setIsCreating(false);
        setProductDetails('');
        setBuyerAddress('');
        fetchOrders(contract, account);
        addToast('success', 'Shipment registered successfully');
      } catch (err) {
        console.error(err);
        addToast('error', 'Error registering shipment.');
        setIsCreating(false);
      }
    }
  };

  const updateStatus = async (orderId, newStatus) => {
    if (contract) {
      try {
        const tx = await contract.updateOrderStatus(orderId, newStatus);
        await tx.wait();
        fetchOrders(contract, account);
      } catch (err) {
        console.error(err);
        addToast('error', 'Error updating shipment status.');
      }
    }
  };

  const logSensorReading = async (orderId) => {
    if (contract) {
        try {
            const temp = parseInt(simTemp);
            const tx = await contract.logIoTReading(orderId, temp, simLocation);
            await tx.wait();
            fetchOrders(contract, account);
        } catch (err) {
            console.error(err);
            addToast('error', 'Error logging sensor data.');
        }
    }
  }

  const cancelOrder = async (orderId) => {
    if (contract) {
      try {
        const tx = await contract.cancelOrder(orderId);
        await tx.wait();
        fetchOrders(contract, account);
      } catch (err) {
        console.error(err);
        addToast('error', 'Error canceling shipment.');
      }
    }
  };

  const copyToClipboard = (text) => {
    navigator.clipboard.writeText(text);
  };

  const getStatusBadge = (status) => {
    switch(status) {
      case 0: return <span className="px-2.5 py-1 bg-gray-100 text-gray-700 rounded-md text-xs font-medium border border-gray-200">Registered</span>;
      case 1: return <span className="px-2.5 py-1 bg-amber-50 text-amber-700 rounded-md text-xs font-medium border border-amber-200">Processing</span>;
      case 2: return <span className="px-2.5 py-1 bg-blue-50 text-blue-700 rounded-md text-xs font-medium border border-blue-200">In Transit</span>;
      case 3: return <span className="px-2.5 py-1 bg-indigo-50 text-indigo-700 rounded-md text-xs font-medium border border-indigo-200">Out for Delivery</span>;
      case 4: return <span className="px-2.5 py-1 bg-emerald-50 text-emerald-700 rounded-md text-xs font-medium border border-emerald-200">Delivered</span>;
      case 5: return <span className="px-2.5 py-1 bg-red-50 text-red-700 rounded-md text-xs font-medium border border-red-200">Canceled</span>;
      default: return null;
    }
  };

  const filteredOrders = orders.filter(o => activeTab === 'all' ? true : o.status < 4);

  return (
    <div className="min-h-screen bg-slate-50 text-slate-900 font-sans selection:bg-teal-100 selection:text-teal-900">
      
      {/* Top Navigation */}
      <nav className="bg-slate-900 border-b border-slate-800 sticky top-0 z-50 shadow-sm">
        <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8">
          <div className="flex justify-between h-16 items-center">
            
            <div className="flex items-center gap-3">
              <div className="bg-teal-500 p-1.5 rounded-lg">
                <Box className="text-white" size={22} strokeWidth={3} />
              </div>
              <span className="text-lg font-bold text-white tracking-tight">
                Covantis<span className="text-slate-400 font-normal">Trust</span>
              </span>
            </div>

            <div className="flex items-center gap-4">
              {account ? (
                <div className="flex items-center gap-3 bg-slate-800 border border-slate-700 px-3 py-1.5 rounded-lg">
                  <div className="flex h-2.5 w-2.5 relative">
                    <span className="animate-ping absolute inline-flex h-full w-full rounded-full bg-emerald-400 opacity-75"></span>
                    <span className="relative inline-flex rounded-full h-2.5 w-2.5 bg-emerald-500"></span>
                  </div>
                  <span className="text-sm font-medium text-slate-300 font-mono">
                    {account.slice(0, 6)}...{account.slice(-4)}
                  </span>
                </div>
              ) : (
                <button 
                  onClick={connectWallet} 
                  className="bg-teal-600 hover:bg-teal-500 text-white text-sm font-medium py-2 px-4 rounded-lg transition-colors flex items-center gap-2"
                >
                  <Wallet size={16} />
                  Connect Wallet
                </button>
              )}
            </div>
          </div>
        </div>
      </nav>

      <main className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-8">
        
        {!account ? (
          <div className="mt-16 bg-white border border-slate-200 rounded-2xl p-12 text-center max-w-2xl mx-auto shadow-sm">
            <div className="mx-auto w-16 h-16 bg-teal-50 text-teal-600 rounded-2xl flex items-center justify-center mb-6">
              <ShieldCheck size={32} />
            </div>
            <h1 className="text-2xl font-bold text-slate-900 mb-3">Supply Chain Logistics</h1>
            <p className="text-slate-500 mb-8 max-w-md mx-auto">
            </p>
            <button 
              onClick={connectWallet} 
              className="bg-slate-900 hover:bg-slate-800 text-white font-medium py-2.5 px-6 rounded-lg transition-colors shadow-sm inline-flex items-center gap-2"
            >
              <Wallet size={18} />
              Connect with MetaMask
            </button>
          </div>
        ) : (
          <div className="grid grid-cols-1 lg:grid-cols-12 gap-8">
            
            {/* Left Column: Register Form & Stats */}
            <div className="lg:col-span-4 space-y-6">
              
              {/* Registration Card */}
              <div className="bg-white border border-slate-200 rounded-xl shadow-sm overflow-hidden">
                <div className="px-5 py-4 border-b border-slate-100 flex items-center justify-between bg-slate-50/50">
                  <h2 className="text-sm font-semibold text-slate-800 flex items-center gap-2">
                    <QrCode size={16} className="text-teal-600" />
                    Register Authenticated Batch
                  </h2>
                </div>
                <div className="p-5">
                  <form onSubmit={createOrder} className="space-y-4">
                    <div>
                      <label className="block text-xs font-medium text-slate-600 mb-1.5 uppercase tracking-wide">Product Payload</label>
                      <input 
                        type="text" 
                        value={productDetails} 
                        onChange={(e) => setProductDetails(e.target.value)} 
                        placeholder="e.g. Pfizer Vaccines Batch rgba(157, 69, 157, 0.27)"
                        className="w-full text-sm rounded-lg border-slate-200 border px-3 py-2.5 focus:ring-2 focus:ring-teal-600/20 focus:border-teal-600 outline-none transition-all placeholder:text-slate-400"
                        required 
                      />
                    </div>
                    <div>
                      <label className="block text-xs font-medium text-slate-600 mb-1.5 uppercase tracking-wide">Recipient Node (0x)</label>
                      <input 
                        type="text" 
                        value={buyerAddress} 
                        onChange={(e) => setBuyerAddress(e.target.value)} 
                        placeholder="meta mask address"
                        className="w-full font-mono text-sm rounded-lg border-slate-200 border px-3 py-2.5 focus:ring-2 focus:ring-teal-600/20 focus:border-teal-600 outline-none transition-all placeholder:text-slate-400"
                        required 
                      />
                    </div>
                    <button 
                      type="submit" 
                      disabled={isCreating}
                      className="w-full bg-[pink] hover:bg-[white] disabled:bg-slate-300 text-white text-sm font-medium py-2.5 rounded-lg transition-colors flex justify-center items-center gap-2 mt-2"
                    >
                      {isCreating ? 'Minting on Chain...' : 'Generate Traceability Record'}
                    </button>
                  </form>
                </div>
              </div>

              {/* Quick Stats */}
              <div className="bg-white border border-slate-200 rounded-xl shadow-sm p-5">
                <h3 className="text-xs font-medium text-slate-500 uppercase tracking-wide mb-4">Hyperledger Overview</h3>
                <div className="grid grid-cols-2 gap-4">
                  <div className="p-3 bg-slate-50 rounded-lg border border-slate-100">
                    <p className="text-xs text-slate-500 mb-1">Handoffs Recorded</p>
                    <p className="text-2xl font-semibold text-slate-900">{orders.length}</p>
                  </div>
                  <div className="p-3 bg-teal-50/50 rounded-lg border border-teal-100">
                    <p className="text-xs text-teal-700 mb-1">Active Smart Contracts</p>
                    <p className="text-2xl font-semibold text-teal-900">{orders.filter(o => o.status < 4).length}</p>
                  </div>
                </div>
              </div>

              {/* Demo IoT Sensor Injector */}
              <div className="bg-white border border-teal-100 rounded-xl shadow-sm overflow-hidden">
                <div className="px-5 py-3 border-b border-teal-100 bg-teal-50/30">
                </div>
                
              </div>

            </div>

            {/* Right Column: Shipment Log */}
            <div className="lg:col-span-8">
              <div className="bg-white border border-slate-200 rounded-xl shadow-sm overflow-hidden min-h-[600px]">
                
                {/* Header & Tabs */}
                <div className="px-6 py-5 border-b border-slate-100 flex flex-col sm:flex-row sm:items-center justify-between gap-4 bg-slate-50/50">
                  <h2 className="text-base font-semibold text-slate-800 flex items-center gap-2">
                    <Activity size={18} className="text-teal-600" />
                    Immutable Logistics Ledger
                  </h2>
                  <div className="flex bg-slate-100 p-1 rounded-lg border border-slate-200">
                    <button 
                      onClick={() => setActiveTab('active')}
                      className={`px-4 py-1.5 text-xs font-medium rounded-md transition-all ${activeTab === 'active' ? 'bg-white text-slate-900 shadow-sm border border-slate-200/50' : 'text-slate-500 hover:text-slate-700'}`}
                    >
                      Active Cargo
                    </button>
                    <button 
                      onClick={() => setActiveTab('all')}
                      className={`px-4 py-1.5 text-xs font-medium rounded-md transition-all ${activeTab === 'all' ? 'bg-white text-slate-900 shadow-sm border border-slate-200/50' : 'text-slate-500 hover:text-slate-700'}`}
                    >
                      Auditable History
                    </button>
                  </div>
                </div>

                {/* List */}
                <div className="divide-y divide-slate-100">
                  {filteredOrders.length === 0 ? (
                    <div className="p-12 text-center text-slate-500">
                      <Package size={48} className="mx-auto text-slate-300 mb-4" strokeWidth={1} />
                      <p className="text-sm font-medium text-slate-900">No verifiable shipments</p>
                      <p className="text-sm mt-1">Register a payload to begin tracking chain of custody.</p>
                    </div>
                  ) : (
                    filteredOrders.map((order) => {
                      const isSeller = order.seller.toLowerCase() === account.toLowerCase();
                      const isBuyer = order.buyer.toLowerCase() === account.toLowerCase();
                      const isCanceled = order.status === 5;
                      const isDelivered = order.status === 4;

                      return (
                        <div key={order.id} className="p-6 hover:bg-slate-50/50 transition-colors">
                          
                          {/* Top row: Status & IDs */}
                          <div className="flex items-center gap-3 mb-4">
                            <span className="font-mono text-xs font-semibold text-slate-700 bg-slate-100 px-2 py-0.5 rounded border border-slate-200">
                              B/L-{order.id.padStart(4, '0')}
                            </span>
                            {getStatusBadge(order.status)}
                            <span className="text-xs text-slate-400 flex items-center gap-1">
                              <Clock size={12} /> {order.updatedAt}
                            </span>
                            
                            {/* Food Safety Trust Signal */}
                            <div className="ml-auto flex items-center gap-1.5 text-xs font-medium text-emerald-600 bg-emerald-50 px-2 py-1 rounded border border-emerald-100">
                              <ShieldCheck size={14} /> Provenance Verified
                            </div>
                          </div>
                          
                          <div className="flex flex-col gap-6">
                            
                            {/* Top: Image, Info & Actions */}
                            <div className="flex flex-col md:flex-row gap-6">
                              {/* Product Image */}
                              <div className="w-32 h-32 shrink-0 bg-white rounded-xl overflow-hidden border border-slate-200 shadow-sm">
                                <img src={order.imageUrl || 'https://via.placeholder.com/150?text=No+Image'} alt="Product" className="w-full h-full object-cover" />
                              </div>
                              
                              {/* Middle: Details & IoT */}
                              <div className="flex-1 flex flex-col justify-between">
                                <div>
                                  <h3 className="text-xl font-bold text-slate-900 mb-1">{order.productDetails}</h3>
                                  <p className="text-sm text-slate-500 mb-3">Ordered by: <span className="font-semibold text-slate-800">{order.customerName}</span> ({order.customerEmail})</p>
                                  
                                  <div className="flex flex-wrap gap-x-6 gap-y-2 text-sm mb-4">
                                    <div className="flex items-center gap-1.5 text-slate-600 bg-slate-50 border border-slate-200 rounded-md px-2 py-1">
                                      <Hash size={14} className="text-slate-400" />
                                      <span className="font-mono text-xs">{order.seller.slice(0, 6)}...{order.seller.slice(-4)}</span>
                                    </div>
                                    <div className="flex items-center justify-center text-slate-300">
                                      <ArrowRight size={16} />
                                    </div>
                                    <div className="flex items-center gap-1.5 text-slate-600 bg-slate-50 border border-slate-200 rounded-md px-2 py-1">
                                      <MapPin size={14} className="text-slate-400" />
                                      <span className="font-mono text-xs">{order.buyer.slice(0, 6)}...{order.buyer.slice(-4)}</span>
                                    </div>
                                  </div>
                                </div>
                                
                                {/* IoT Logs inline */}
                                <div className={`px-3 py-2 rounded-md border flex flex-col sm:flex-row sm:items-center justify-between gap-3 ${order.conditionViolated ? 'bg-red-50 border-red-200' : 'bg-slate-50 border-slate-200'}`}>
                                  <div className="flex items-center gap-2">
                                    <Thermometer size={14} className="text-slate-500" />
                                    <span className="text-xs font-bold text-slate-600 uppercase tracking-wider">IoT Sensor Data</span>
                                  </div>
                                  <div className="flex-1 flex flex-wrap gap-2">
                                    {order.readings && order.readings.length > 0 ? (
                                        order.readings.map((r, i) => (
                                            <div key={i} className={`text-[10px] font-mono px-1.5 py-0.5 rounded border ${r.temp > 8 || r.temp < 2 ? 'bg-red-100 border-red-300 text-red-800' : 'bg-white border-slate-200 text-slate-600'}`}>
                                                {r.temp}°C @ {r.loc} ({r.time})
                                            </div>
                                        ))
                                    ) : (
                                        <span className="text-[11px] text-slate-400 italic">No telemetry yet.</span>
                                    )}
                                  </div>
                                  <div>
                                    {order.conditionViolated ? (
                                      <span className="text-[10px] font-bold text-red-600 bg-red-100 px-2 py-1 rounded border border-red-200 flex items-center gap-1"><ShieldAlert size={12}/> SPOILED</span>
                                    ) : (
                                      <span className="text-[10px] font-bold text-emerald-700 bg-emerald-100 px-2 py-1 rounded border border-emerald-200 flex items-center gap-1"><CheckCircle2 size={12}/> COMPLIANT</span>
                                    )}
                                  </div>
                                </div>
                              </div>

                              {/* Right: Actions */}
                              <div className="w-full md:w-48 shrink-0 flex flex-col gap-2">
                                {order.status < 4 && !isCanceled && (
                                  <>
                                    {isSeller && (
                                      <button 
                                        onClick={() => updateStatus(order.id, order.status + 1)}
                                        className="w-full bg-amber-400 hover:bg-amber-500 text-amber-950 text-sm font-semibold py-2 px-3 rounded-lg shadow-sm transition-colors flex justify-center items-center gap-1"
                                      >
                                        Update Status <ChevronRight size={16} />
                                      </button>
                                    )}
                                    {isSeller && (
                                      <button 
                                        onClick={() => logSensorReading(order.id)}
                                        className="w-full bg-white border border-teal-200 hover:bg-teal-50 text-teal-700 text-xs font-medium py-1.5 px-3 rounded-lg shadow-sm transition-colors flex justify-center items-center gap-1"
                                      >
                                        <Thermometer size={14}/> Push Sensor Data
                                      </button>
                                    )}
                                    {(isBuyer || isSeller) && order.status < 2 && (
                                      <button 
                                        onClick={() => cancelOrder(order.id)}
                                        className="w-full mt-auto py-1.5 bg-white border border-slate-200 hover:bg-red-50 hover:text-red-700 hover:border-red-200 text-slate-500 text-xs font-medium rounded-lg shadow-sm transition-colors"
                                      >
                                        Cancel Shipment
                                      </button>
                                    )}
                                  </>
                                )}
                              </div>
                            </div>

                            {/* Horizontal Progress Tracker (Amazon style) */}
                            {!isCanceled && (
                              <div className="w-full pt-6 pb-2 border-t border-slate-100">
                                <h4 className="text-lg font-bold text-slate-800 mb-6">
                                  {order.status === 4 ? "Delivered" : "Arriving soon"}
                                </h4>
                                <div className="relative px-4">
                                  {/* Line */}
                                  <div className="absolute top-2 left-8 right-8 h-1 bg-slate-200 rounded-full"></div>
                                  <div 
                                    className={`absolute top-2 left-8 h-1 rounded-full transition-all duration-1000 ${order.conditionViolated ? 'bg-red-500' : 'bg-emerald-500'}`}
                                    style={{ width: `calc(${(order.status / 4) * 100}% - 2rem)` }}
                                  ></div>
                                  
                                  {/* Steps */}
                                  <div className="relative flex justify-between z-10">
                                    {["Ordered", "Processing", "Shipped", "Out for delivery", "Delivered"].map((label, index) => {
                                      const isCompleted = order.status >= index;
                                      const isCurrent = order.status === index;
                                      return (
                                        <div key={label} className="flex flex-col items-center w-16">
                                          <div className={`w-5 h-5 rounded-full border-[3px] shadow-sm transition-colors bg-white ${isCompleted ? (order.conditionViolated ? 'border-red-500' : 'border-emerald-500') : 'border-slate-300'}`}>
                                            {isCompleted && !order.conditionViolated && (
                                              <div className="w-full h-full bg-emerald-500 rounded-full scale-50"></div>
                                            )}
                                            {isCompleted && order.conditionViolated && (
                                              <div className="w-full h-full bg-red-500 rounded-full scale-50"></div>
                                            )}
                                          </div>
                                          <span className={`text-[11px] font-medium mt-2 text-center leading-tight ${isCompleted ? 'text-slate-900' : 'text-slate-400'} ${isCurrent ? 'font-bold' : ''}`}>
                                            {label}
                                          </span>
                                        </div>
                                      );
                                    })}
                                  </div>
                                </div>
                              </div>
                            )}
                          </div>  
                              <div className="mt-auto pt-4 text-[10px] text-slate-400 font-mono flex items-center gap-1 w-full justify-end cursor-pointer hover:text-teal-600">
                                Verify on Explorer <ExternalLink size={10} />
                              </div>
                            </div>
                            
                          </div>
                        </div>
                      )
                    })
                  )}
                </div>

              </div>
            </div>

          </div>
        )}
      </main>
    </div>
  );
}
