## 📦 Delivery Split Classification Macro

- Developed to identify and classify delivery splits based on delivery number patterns  
- Uses prefix logic:
  - **986******* → Initial delivery**
  - **801–899******* → Split deliveries**
  
- Groups related deliveries using:
  - Ship-to party  
  - Route  
  - Planned delivery date  
  - Delivery suffix (last 7 digits)

- Classifies each delivery into defined scenarios:
  - Full delivery (no split)  
  - Single split  
  - Multiple splits  
  - Same-day split cases  

- Outputs results into:
  - **Column AI → Primary classification**
  - **Column AJ → Same-day split logic**

- Helps to:
  - Identify delivery split patterns  
  - Reduce manual analysis  
  - Highlight potential cost-impact scenarios
 
  ## 💻 VBA Code
- [Classify_Final.bas](./Classify_Final.bas)
