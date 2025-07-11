import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import os
df=pd.read_excel("Superstore.xlsx")

st.set_page_config(page_title="Sales Dashboard", layout="wide")
st.title("Sales Dashboard & Revenue Analysis System")

uploaded_file = st.file_uploader("Upload Superstore XLSX File", type=["xlsx"])

if uploaded_file is not None:
    df = pd.read_excel(uploaded_file)

    # 🧽 2. Data Cleaning
    # - Convert 'Order Date' column to datetime
    # - Remove duplicates
    # - Handle missing values (drop or fill)
    # - Clean column formatting (trim spaces, lowercase headers)
    
    df['Order Date'] = pd.to_datetime(df['Order Date'],errors='coerce')
    df.dropna(subset=['Order Date'],inplace=True)
    df.drop_duplicates(inplace=True)
    df.fillna(0,inplace=True)
    df.columns = [col.strip().lower().replace(' ','_')for col in df.columns]
    
    #Create month and quarter columns
    df['month']=df['order_date'].dt.to_period('M')
    df['quarter']=df['order_date'].dt.to_period('Q')

    df

    #  Q1. What are the total sales every month or quarter?
    # - ✅ Use: groupby + sum() on Month/Quarter
    # - ✅ Visual: *Line Plot*
    
    #Setting style
    sns.set_theme(style="whitegrid")
    os.makedirs("images",exist_ok=True)
    
    monthly_sales = df.groupby('month')['sales'].sum()
    monthly_sales.plot(kind='line',marker='o',figsize=(10,5),title="Monthly sales trend")
    plt.xlabel("Month")
    plt.ylabel("Sales")
    plt.tight_layout()
    plt.savefig("images/monthly_sales_trend.png")
    st.pyplot(plt.gcf())

    
    # Q2. Which 10 products brought the most revenue?
    # - ✅ Use: groupby('Product') + sort by 'Sales'
    # - ✅ Visual: *Bar Chart*
    
    top_products =  df.groupby('product_name')['sales'].sum().sort_values(ascending=False).head(10)
    top_products.plot(kind='bar',color='skyblue',figsize=(10,5),title="Top 10 Products by Revenue")
    plt.ylabel("Total Sales")
    plt.savefig("images/top_products.png")
    st.pyplot(plt.gcf())

    
    # Q3. Which region generated the highest sales and profit?
    # - ✅ Use: groupby('Region')
    # - ✅ Visual: *Bar Graph / Pie Chart*
    
    region_data = df.groupby('region')[['sales','profit']].sum().sort_values(by='sales',ascending=False)
    region_data.plot(kind='bar',figsize=(8,5),title="Sales & Profit by Region")
    plt.ylabel("Amount")
    plt.xlabel("Region")
    plt.tight_layout()
    plt.savefig("images/region_sales_profit.png")
    st.pyplot(plt.gcf())

    
    # Q4. Which product category is underperforming in any region?
    # - ✅ Use: pivot_table(Category vs Region)
    # - ✅ Visual: *Heatmap*
    
    pivot = pd.pivot_table(df,values='profit',index='category',columns='region',aggfunc='sum')
    plt.figure(figsize=(8,5))
    sns.heatmap(pivot,annot=True,fmt=".0f",cmap="YlOrRd")
    plt.title("Profit by category and region")
    plt.tight_layout()
    plt.savefig("images/category_region_heatmap.png")
    st.pyplot(plt.gcf())

    
    # Q5. What is the relationship between sales and profit?
    # - ✅ Use: scatter plot
    # - ✅ Visual: Seaborn.scatterplot()
    
    plt.figure(figsize=(8,5))
    sns.scatterplot(data=df,x='sales',y='profit',hue='category',alpha=0.7)
    plt.title("Sales vs Profit")
    plt.tight_layout()
    plt.savefig("images/sales_vs_profit.png")
    st.pyplot(plt.gcf())

    
    # Q6. How many orders were made per month?
    # - ✅ Use: value_counts() or groupby('Order Date')
    # - ✅ Visual: *Bar/Line Chart*
    
    monthly_orders=df['month'].value_counts().sort_index()
    monthly_orders.plot(kind='bar',figsize=(10,5),title="Orders per Month")
    plt.ylabel("Order Count")
    plt.xlabel("Month")
    plt.tight_layout()
    plt.savefig("images/orders_per_month.png")
    st.pyplot(plt.gcf())

    
    # #### 🟡 Q7. Which city or state has maximum orders with low profit?
    # - ✅ Use: groupby('City') and compare Sales vs Profit
    # - ✅ Visual: *Bar or Lollipop Chart*
    
    city_data = df.groupby('city')[['sales','profit']].sum()
    low_profit_cities=city_data[city_data['profit']<1000].sort_values(by='sales',ascending=False).head(10)
    low_profit_cities['sales'].plot(kind='bar',color='orange',figsize=(10,5),title="Top cities with low Profit")
    plt.ylabel("Sales")
    plt.tight_layout()
    plt.savefig("images/low_profitt_cities.png")
    st.pyplot(plt.gcf())

    
    # #### 🟡 Q8. Which customer segment or shipping mode is most used?
    # - ✅ Use: groupby('Segment') or groupby('Ship Mode')
    # - ✅ Visual: *Pie Chart or Countplot*
    
    fig,axes=plt.subplots(1,2,figsize=(12,5))
    sns.countplot(data=df,x='segment',ax=axes[0])
    axes[0].set_title("Customer Segment Distribution")
    sns.countplot(data=df,x='ship_mode',ax=axes[1])
    axes[1].set_title("Shipping Mode Usage")
    plt.tight_layout()
    plt.savefig("images/segment_shipmode.png")
    st.pyplot(plt.gcf())


else:
    st.info("Please upload a Superstore XLSX file to begin analysis.")
