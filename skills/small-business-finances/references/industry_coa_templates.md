# Industry Chart of Accounts & Direct Cost Label Templates

Use these templates when initializing `classifications.json` so the **Summary P&L** speaks the small business owner's everyday language while preserving standard accounting structure.

## 1. Personal Care, Studios, Salons & Clinics (`business_type: "service"`)
- **`direct_cost_label`:** `"Cost of Providing Services"`
- **Typical `Direct costs` categories:**
  - Service supplies & consumables (lashes, skincare, medical/treatment disposables)
  - Service contractors & commission technicians
  - Client packaging & treatment kits
  - Outbound client delivery
- **Typical `Operating expenses` mappings:**
  - `People & Staff`: Salon/clinic staff salaries, welfare, training
  - `Power, Fuel & Maintenance`: Electricity, generator fuel, equipment servicing
  - `Marketing & Promotion`: Social media ads, influencer PR packages, event sponsorships
  - `Premises, Cleaning & Security`: Studio rent/space fee, cleaning, water, security
  - `Admin, Comms & Operations`: Booking phone/data, studio refreshments, local transport

## 2. Retail, Boutiques & E-Commerce (`business_type: "retail"`)
- **`direct_cost_label`:** `"Cost of Goods Sold (COGS)"`
- **Typical `Direct costs` categories:**
  - Wholesale inventory & stock purchases
  - Inbound freight, customs & duties
  - Order packaging, boxes & labels
  - Customer shipping & courier fulfillment
  - Payment gateway transaction fees tied directly to orders (or keep in `Bank fees`)

## 3. Bakeries, Restaurants, Catering & Hospitality (`business_type: "food"`)
- **`direct_cost_label`:** `"Cost of Food, Ingredients & Packaging"`
- **Typical `Direct costs` categories:**
  - Raw ingredients, produce & beverages
  - Takeaway containers, boxes & disposables
  - Event kitchen casual labor / delivery riders

## 4. Agencies, Freelancers & Consultancies (`business_type: "agency"`)
- **`direct_cost_label`:** `"Direct Project & Subcontractor Costs"`
- **Typical `Direct costs` categories:**
  - Subcontractors & freelance specialists billed to client work
  - Client-specific production, media buy pass-throughs, or licensing costs
