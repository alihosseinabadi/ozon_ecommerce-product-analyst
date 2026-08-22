import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime
import warnings
import os
warnings.filterwarnings('ignore')

# ============================================
# PAGE CONFIGURATION
# ============================================
st.set_page_config(
    page_title="Ozon E-Commerce Analytics",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ============================================
# CUSTOM CSS
# ============================================
st.markdown("""
<style>
    .header {
        background: linear-gradient(135deg, #1a1a2e 0%, #16213e 100%);
        padding: 2rem;
        border-radius: 10px;
        margin-bottom: 2rem;
        color: white;
    }
    .header h1 {
        font-size: 2.5rem;
        font-weight: 700;
        margin: 0;
    }
    .header p {
        color: #a8d8ea;
        margin: 0.5rem 0 0 0;
        font-size: 1.1rem;
    }
    .metric-card {
        background: #ffffff;
        padding: 1.2rem;
        border-radius: 10px;
        box-shadow: 0 2px 8px rgba(0,0,0,0.08);
        border-left: 4px solid #1E88E5;
        margin: 0.25rem 0;
    }
    .metric-value {
        font-size: 1.8rem;
        font-weight: 700;
        color: #1a1a2e;
    }
    .metric-label {
        font-size: 0.85rem;
        color: #666;
        text-transform: uppercase;
        letter-spacing: 0.5px;
    }
    .insight-box {
        background: #f8f9fa;
        padding: 1rem;
        border-radius: 8px;
        border-left: 4px solid #1E88E5;
        margin: 0.5rem 0;
    }
    .footer {
        text-align: center;
        padding: 1.5rem;
        color: #666;
        border-top: 1px solid #e0e0e0;
        margin-top: 2rem;
    }
</style>
""", unsafe_allow_html=True)

# ============================================
# DATA LOADING FUNCTION
# ============================================
@st.cache_data
def load_data():
    """Load Ozon product data from Excel file with Russian columns"""
    
    df = None
    
    # Try to load the Excel file
    try:
        df = pd.read_excel('2021-07-12_opendata_datasetfavorites_2021-06-12_2021-07-11.xlsx', engine='openpyxl')
        st.success("Loaded Excel file successfully!")
    except Exception as e:
        st.warning(f"Could not load Excel: {e}")
    
    # If Excel failed, try CSV
    if df is None:
        try:
            df = pd.read_csv('ozon_clean_data.csv', encoding='utf-8-sig')
            st.success("Loaded CSV file successfully!")
        except:
            pass
    
    # If still None, try the text file
    if df is None:
        try:
            df = pd.read_csv('Название товараБрендСсылка на товарКатегория 1.txt', sep='\t', encoding='utf-8')
            st.success("Loaded TXT file successfully!")
        except:
            pass
    
    if df is None:
        st.error("No data file found. Please check the file.")
        return None
    
    # Show original columns
    st.info(f"Original columns: {list(df.columns)}")
    
    # Russian to English column mapping
    russian_columns = {
        'Название товара': 'product_name',
        'Бренд': 'brand',
        'Ссылка на товар': 'url',
        'Категория 1 уровня': 'category_level_1',
        'Категория 2 уровня': 'category_level_2',
        'Категория 3 уровня': 'category_level_3',
        'Категория 4 уровня': 'category_level_4',
        'Количество добавлений в избранное, 30 дней': 'favorites_30d',
        'Количество добавлений в избранное, все время': 'favorites_all',
        'Последнее появление в наличии': 'last_in_stock'
    }
    
    # Rename columns that exist
    rename_dict = {k: v for k, v in russian_columns.items() if k in df.columns}
    if rename_dict:
        df = df.rename(columns=rename_dict)
        st.info(f"Renamed columns: {list(rename_dict.values())}")
    
    # Ensure all required columns exist
    if 'product_name' not in df.columns:
        for col in df.columns:
            if 'название' in str(col).lower() or 'product' in str(col).lower() or 'name' in str(col).lower():
                df['product_name'] = df[col]
                break
        if 'product_name' not in df.columns:
            df['product_name'] = 'Unknown'
    
    if 'brand' not in df.columns:
        for col in df.columns:
            if 'бренд' in str(col).lower() or 'brand' in str(col).lower():
                df['brand'] = df[col]
                break
        if 'brand' not in df.columns:
            df['brand'] = 'Unknown'
    
    if 'category_level_1' not in df.columns:
        for col in df.columns:
            if 'категория' in str(col).lower() or 'category' in str(col).lower():
                df['category_level_1'] = df[col]
                break
        if 'category_level_1' not in df.columns:
            df['category_level_1'] = 'Uncategorized'
    
    if 'favorites_all' not in df.columns:
        for col in df.columns:
            if 'избранное' in str(col).lower() or 'favorite' in str(col).lower():
                df['favorites_all'] = df[col]
                break
        if 'favorites_all' not in df.columns:
            df['favorites_all'] = 0
    
    if 'favorites_30d' not in df.columns:
        for col in df.columns:
            if '30' in str(col) and ('избранное' in str(col).lower() or 'favorite' in str(col).lower()):
                df['favorites_30d'] = df[col]
                break
        if 'favorites_30d' not in df.columns:
            df['favorites_30d'] = 0
    
    # Clean data - replace NaN with appropriate values
    df['product_name'] = df['product_name'].fillna('Unknown')
    df['brand'] = df['brand'].fillna('Unknown')
    df['category_level_1'] = df['category_level_1'].fillna('Uncategorized')
    
    # Convert to numeric, handling errors
    df['favorites_30d'] = pd.to_numeric(df['favorites_30d'], errors='coerce').fillna(0).astype(int)
    df['favorites_all'] = pd.to_numeric(df['favorites_all'], errors='coerce').fillna(0).astype(int)
    
    # Calculate derived metrics
    df['recent_ratio'] = (df['favorites_30d'] / df['favorites_all'] * 100).fillna(0)
    df['recent_ratio'] = df['recent_ratio'].replace([np.inf, -np.inf], 0)
    
    # Remove rows with missing critical data
    df = df[df['product_name'] != 'Unknown']
    
    st.success(f"Data loaded successfully! {len(df)} rows, {len(df.columns)} columns")
    
    return df

# ============================================
# SAFE CHART FUNCTION
# ============================================
def safe_chart(fig, message="No data available"):
    """Safely display a chart with error handling"""
    if fig is None:
        st.info(message)
        return False
    try:
        st.plotly_chart(fig, use_container_width=True)
        return True
    except Exception as e:
        st.info(f"{message}: {str(e)}")
        return False

def create_safe_bar(data, x, y, **kwargs):
    """Create a bar chart safely"""
    if data is None or len(data) == 0:
        return None
    try:
        return px.bar(data, x=x, y=y, **kwargs)
    except:
        return None

def create_safe_scatter(data, x, y, **kwargs):
    """Create a scatter chart safely"""
    if data is None or len(data) == 0:
        return None
    try:
        return px.scatter(data, x=x, y=y, **kwargs)
    except:
        return None

def create_safe_histogram(data, x, **kwargs):
    """Create a histogram safely"""
    if data is None or len(data) == 0:
        return None
    try:
        return px.histogram(data, x=x, **kwargs)
    except:
        return None

# ============================================
# LOAD DATA
# ============================================
df = load_data()

if df is None:
    st.stop()

# ============================================
# SIDEBAR
# ============================================
with st.sidebar:
    st.markdown("### Ozon Analytics")
    st.markdown("---")
    
    st.markdown("### Dashboard Controls")
    
    # Category filter
    categories_list = df['category_level_1'].dropna().unique().tolist()
    categories_list = [str(cat) for cat in categories_list if str(cat) != 'nan' and str(cat) != 'Unknown' and str(cat) != 'Uncategorized']
    if categories_list:
        categories = ['All'] + sorted(categories_list)
    else:
        categories = ['All']
    selected_category = st.selectbox("Select Category", categories)
    
    # Brand filter
    brands_list = df['brand'].dropna().unique().tolist()
    brands_list = [str(b) for b in brands_list if str(b) != 'nan' and str(b) != 'Unknown']
    if brands_list:
        brands = ['All'] + sorted(brands_list)
    else:
        brands = ['All']
    selected_brand = st.selectbox("Select Brand", brands)
    
    # Min favorites filter
    max_fav = int(df['favorites_all'].max()) if len(df) > 0 else 0
    min_favorites = st.slider("Minimum Favorites (All Time)", 0, max_fav, 0)
    
    st.markdown("---")
    st.markdown("### Data Info")
    st.write(f"Total Products: {len(df):,}")
    st.write(f"Total Brands: {len(brands_list):,}")
    st.write(f"Total Categories: {len(categories_list):,}")

# ============================================
# FILTER DATA
# ============================================
filtered_df = df.copy()

if selected_category != 'All':
    filtered_df = filtered_df[filtered_df['category_level_1'] == selected_category]

if selected_brand != 'All':
    filtered_df = filtered_df[filtered_df['brand'] == selected_brand]

filtered_df = filtered_df[filtered_df['favorites_all'] >= min_favorites]

# ============================================
# HEADER
# ============================================
st.markdown("""
<div class="header">
    <h1>Ozon E-Commerce Analytics Dashboard</h1>
    <p>Product performance analysis and business intelligence</p>
</div>
""", unsafe_allow_html=True)

# ============================================
# KEY METRICS
# ============================================
col1, col2, col3, col4, col5 = st.columns(5)

with col1:
    st.markdown(f"""
    <div class="metric-card">
        <div class="metric-label">Total Products</div>
        <div class="metric-value">{len(filtered_df):,}</div>
    </div>
    """, unsafe_allow_html=True)

with col2:
    total_fav = filtered_df['favorites_all'].sum() if len(filtered_df) > 0 else 0
    st.markdown(f"""
    <div class="metric-card">
        <div class="metric-label">Total Favorites (All Time)</div>
        <div class="metric-value">{total_fav:,.0f}</div>
    </div>
    """, unsafe_allow_html=True)

with col3:
    total_fav_30d = filtered_df['favorites_30d'].sum() if len(filtered_df) > 0 else 0
    st.markdown(f"""
    <div class="metric-card">
        <div class="metric-label">Total Favorites (30 Days)</div>
        <div class="metric-value">{total_fav_30d:,.0f}</div>
    </div>
    """, unsafe_allow_html=True)

with col4:
    avg_fav = filtered_df['favorites_all'].mean() if len(filtered_df) > 0 else 0
    st.markdown(f"""
    <div class="metric-card">
        <div class="metric-label">Avg Favorites per Product</div>
        <div class="metric-value">{avg_fav:.1f}</div>
    </div>
    """, unsafe_allow_html=True)

with col5:
    brands_count = filtered_df['brand'].nunique() if len(filtered_df) > 0 else 0
    st.markdown(f"""
    <div class="metric-card">
        <div class="metric-label">Unique Brands</div>
        <div class="metric-value">{brands_count:,}</div>
    </div>
    """, unsafe_allow_html=True)

st.markdown("---")

# ============================================
# TABS
# ============================================
tab1, tab2, tab3, tab4 = st.tabs(["Overview", "Category Analysis", "Brand Analysis", "Trends"])

# ============================================
# TAB 1: OVERVIEW
# ============================================
with tab1:
    if len(filtered_df) == 0:
        st.warning("No data available with current filters. Please adjust your filters.")
    else:
        col1, col2 = st.columns(2)
        
        with col1:
            st.subheader("Top Categories by Product Count")
            category_counts = filtered_df['category_level_1'].value_counts().head(10)
            if len(category_counts) > 0:
                fig = px.bar(
                    x=category_counts.values,
                    y=category_counts.index,
                    orientation='h',
                    color=category_counts.values,
                    color_continuous_scale='Blues',
                    labels={'x': 'Number of Products', 'y': 'Category'}
                )
                fig.update_layout(height=400, margin=dict(l=0, r=0, t=20, b=0))
                st.plotly_chart(fig, use_container_width=True)
            else:
                st.info("No category data available")
        
        with col2:
            st.subheader("Top Brands by Product Count")
            brand_counts = filtered_df['brand'].value_counts().head(10)
            if len(brand_counts) > 0:
                fig = px.bar(
                    x=brand_counts.values,
                    y=brand_counts.index,
                    orientation='h',
                    color=brand_counts.values,
                    color_continuous_scale='Oranges',
                    labels={'x': 'Number of Products', 'y': 'Brand'}
                )
                fig.update_layout(height=400, margin=dict(l=0, r=0, t=20, b=0))
                st.plotly_chart(fig, use_container_width=True)
            else:
                st.info("No brand data available")
        
        col1, col2 = st.columns(2)
        
        with col1:
            st.subheader("Top 10 Products by Favorites")
            top_products = filtered_df.nlargest(10, 'favorites_all')[['product_name', 'brand', 'favorites_all']]
            if len(top_products) > 0:
                fig = px.bar(
                    top_products,
                    x='favorites_all',
                    y='product_name',
                    orientation='h',
                    color='favorites_all',
                    color_continuous_scale='Greens',
                    labels={'favorites_all': 'Total Favorites', 'product_name': 'Product'}
                )
                fig.update_layout(height=400, margin=dict(l=0, r=0, t=20, b=0))
                st.plotly_chart(fig, use_container_width=True)
            else:
                st.info("No products with favorites")
        
        with col2:
            st.subheader("Favorites Distribution")
            try:
                if len(filtered_df) > 1:
                    fig = px.histogram(
                        filtered_df,
                        x='favorites_all',
                        nbins=50,
                        color_discrete_sequence=['purple'],
                        labels={'favorites_all': 'Total Favorites', 'count': 'Number of Products'}
                    )
                    fig.update_layout(height=400, margin=dict(l=0, r=0, t=20, b=0))
                    fig.update_xaxis(type="log")
                    st.plotly_chart(fig, use_container_width=True)
                else:
                    st.info("Not enough data for distribution chart")
            except Exception as e:
                st.info(f"Could not create distribution chart: {str(e)[:50]}")

# ============================================
# TAB 2: CATEGORY ANALYSIS
# ============================================
with tab2:
    if len(filtered_df) == 0:
        st.warning("No data available with current filters. Please adjust your filters.")
    else:
        st.subheader("Category Performance Metrics")
        
        category_stats = filtered_df.groupby('category_level_1').agg({
            'favorites_all': ['sum', 'mean'],
            'favorites_30d': ['sum'],
            'product_name': 'count'
        }).round(2)
        
        category_stats.columns = ['total_favorites', 'avg_favorites', 'favorites_30d', 'product_count']
        category_stats['favorites_per_product'] = category_stats['total_favorites'] / category_stats['product_count']
        category_stats['recent_ratio'] = (category_stats['favorites_30d'] / category_stats['total_favorites']) * 100
        category_stats['recent_ratio'] = category_stats['recent_ratio'].fillna(0)
        category_stats = category_stats.sort_values('total_favorites', ascending=False)
        
        if len(category_stats) > 0:
            col1, col2 = st.columns(2)
            
            with col1:
                fig = px.bar(
                    category_stats.head(10),
                    x='total_favorites',
                    y=category_stats.head(10).index,
                    orientation='h',
                    color='avg_favorites',
                    color_continuous_scale='Viridis',
                    labels={'total_favorites': 'Total Favorites', 'category_level_1': 'Category'}
                )
                fig.update_layout(height=400, title="Total Favorites by Category")
                st.plotly_chart(fig, use_container_width=True)
            
            with col2:
                fig = px.scatter(
                    category_stats,
                    x='product_count',
                    y='avg_favorites',
                    size='total_favorites',
                    color='recent_ratio',
                    color_continuous_scale='RdYlGn',
                    labels={'product_count': 'Number of Products', 'avg_favorites': 'Avg Favorites per Product'},
                    hover_data=['total_favorites']
                )
                fig.update_layout(height=400, title="Category Performance Matrix")
                st.plotly_chart(fig, use_container_width=True)
            
            with st.expander("View Category Details"):
                display_df = category_stats.copy()
                display_df['total_favorites'] = display_df['total_favorites'].apply(lambda x: f"{x:,.0f}")
                display_df['avg_favorites'] = display_df['avg_favorites'].apply(lambda x: f"{x:.1f}")
                display_df['favorites_30d'] = display_df['favorites_30d'].apply(lambda x: f"{x:,.0f}")
                display_df['product_count'] = display_df['product_count'].apply(lambda x: f"{x:,.0f}")
                display_df['favorites_per_product'] = display_df['favorites_per_product'].apply(lambda x: f"{x:.1f}")
                display_df['recent_ratio'] = display_df['recent_ratio'].apply(lambda x: f"{x:.1f}%")
                st.dataframe(display_df, use_container_width=True)
        else:
            st.info("No category data available")

# ============================================
# TAB 3: BRAND ANALYSIS
# ============================================
with tab3:
    if len(filtered_df) == 0:
        st.warning("No data available with current filters. Please adjust your filters.")
    else:
        st.subheader("Brand Performance Analysis")
        
        brand_stats = filtered_df.groupby('brand').agg({
            'favorites_all': ['sum', 'mean'],
            'favorites_30d': ['sum'],
            'product_name': 'count'
        }).round(2)
        
        brand_stats.columns = ['total_favorites', 'avg_favorites', 'favorites_30d', 'product_count']
        brand_stats['favorites_per_product'] = brand_stats['total_favorites'] / brand_stats['product_count']
        brand_stats = brand_stats[brand_stats['product_count'] >= 5]
        brand_stats = brand_stats.sort_values('total_favorites', ascending=False)
        
        if len(brand_stats) == 0:
            st.info("Need at least 5 products per brand for brand analysis")
        else:
            col1, col2 = st.columns(2)
            
            with col1:
                fig = px.bar(
                    brand_stats.head(10),
                    x='total_favorites',
                    y=brand_stats.head(10).index,
                    orientation='h',
                    color='avg_favorites',
                    color_continuous_scale='Reds',
                    labels={'total_favorites': 'Total Favorites', 'brand': 'Brand'}
                )
                fig.update_layout(height=400, title="Top Brands by Total Favorites")
                st.plotly_chart(fig, use_container_width=True)
            
            with col2:
                fig = px.scatter(
                    brand_stats,
                    x='product_count',
                    y='avg_favorites',
                    size='total_favorites',
                    color='favorites_per_product',
                    color_continuous_scale='Blues',
                    labels={'product_count': 'Number of Products', 'avg_favorites': 'Avg Favorites per Product'},
                    hover_data=['total_favorites']
                )
                fig.update_layout(height=400, title="Brand Performance Matrix")
                st.plotly_chart(fig, use_container_width=True)
            
            st.subheader("Most Efficient Brands (Highest Avg Favorites per Product)")
            efficient_brands = brand_stats.nlargest(10, 'avg_favorites')[['product_count', 'avg_favorites', 'total_favorites']]
            
            if len(efficient_brands) > 0:
                fig = px.bar(
                    efficient_brands,
                    x=efficient_brands.index,
                    y='avg_favorites',
                    color='product_count',
                    color_continuous_scale='Greens',
                    labels={'brand': 'Brand', 'avg_favorites': 'Avg Favorites per Product'}
                )
                fig.update_layout(height=350)
                st.plotly_chart(fig, use_container_width=True)

# ============================================
# TAB 4: TRENDS
# ============================================
with tab4:
    if len(filtered_df) == 0:
        st.warning("No data available with current filters. Please adjust your filters.")
    else:
        st.subheader("Recent vs All Time Performance")
        
        col1, col2 = st.columns(2)
        
        with col1:
            if len(filtered_df) > 1:
                sample_df = filtered_df.sample(min(500, len(filtered_df)))
                fig = px.scatter(
                    sample_df,
                    x='favorites_all',
                    y='favorites_30d',
                    color='category_level_1',
                    size='favorites_30d',
                    hover_data=['product_name', 'brand'],
                    labels={'favorites_all': 'All Time Favorites', 'favorites_30d': '30-Day Favorites'},
                    title="Recent vs All Time Favorites"
                )
                fig.update_layout(height=400)
                fig.update_xaxes(type="log")
                fig.update_yaxis(type="log")
                st.plotly_chart(fig, use_container_width=True)
            else:
                st.info("Not enough data for scatter plot")
        
        with col2:
            trending = filtered_df.nlargest(10, 'favorites_30d')[['product_name', 'brand', 'favorites_30d', 'favorites_all']]
            if len(trending) > 0:
                trending['recent_ratio'] = (trending['favorites_30d'] / trending['favorites_all'] * 100).round(1)
                
                fig = px.bar(
                    trending,
                    x='favorites_30d',
                    y='product_name',
                    orientation='h',
                    color='recent_ratio',
                    color_continuous_scale='Reds',
                    labels={'favorites_30d': '30-Day Favorites', 'product_name': 'Product'},
                    title="Top Trending Products (Last 30 Days)"
                )
                fig.update_layout(height=400)
                st.plotly_chart(fig, use_container_width=True)
            else:
                st.info("No trending products")
        
        # Key Insights
        st.subheader("Key Business Insights")
        
        col1, col2, col3 = st.columns(3)
        
        with col1:
            top_cat = filtered_df['category_level_1'].value_counts().index[0]
            top_cat_count = filtered_df['category_level_1'].value_counts().iloc[0]
            st.markdown(f"""
            <div class="insight-box">
                <h4>Top Category</h4>
                <p><strong>{top_cat}</strong></p>
                <p>{top_cat_count:,} products ({top_cat_count/len(filtered_df)*100:.1f}%)</p>
            </div>
            """, unsafe_allow_html=True)
        
        with col2:
            top_brand = filtered_df['brand'].value_counts().index[0]
            top_brand_count = filtered_df['brand'].value_counts().iloc[0]
            st.markdown(f"""
            <div class="insight-box">
                <h4>Top Brand</h4>
                <p><strong>{top_brand}</strong></p>
                <p>{top_brand_count:,} products</p>
            </div>
            """, unsafe_allow_html=True)
        
        with col3:
            top_product = filtered_df.loc[filtered_df['favorites_all'].idxmax()]
            st.markdown(f"""
            <div class="insight-box">
                <h4>Most Popular Product</h4>
                <p><strong>{top_product['product_name'][:40]}...</strong></p>
                <p>{top_product['favorites_all']:,} favorites</p>
                <p>Brand: {top_product['brand']}</p>
            </div>
            """, unsafe_allow_html=True)
        
        # Recommendations
        st.subheader("Business Recommendations")
        
        recommendations = []
        
        if len(category_stats) > 0:
            top_cat_rec = category_stats.index[0]
            recommendations.append({
                'priority': 'High',
                'area': 'Category Focus',
                'action': f'Focus on {top_cat_rec} category',
                'reason': f'Largest category with {category_stats.loc[top_cat_rec, "product_count"]:.0f} products'
            })
        
        if 'brand_stats' in locals() and len(brand_stats) > 0:
            top_brand_rec = brand_stats.index[0]
            recommendations.append({
                'priority': 'High',
                'area': 'Brand Strategy',
                'action': f'Partner with {top_brand_rec}',
                'reason': f'Top brand with {brand_stats.loc[top_brand_rec, "product_count"]:.0f} products'
            })
        
        if len(category_stats) > 0:
            opportunity_cat = category_stats.nlargest(5, 'avg_favorites').index[0]
            if category_stats.loc[opportunity_cat, 'product_count'] < 100:
                recommendations.append({
                    'priority': 'Medium',
                    'area': 'Category Expansion',
                    'action': f'Expand products in {opportunity_cat}',
                    'reason': f'High avg favorites ({category_stats.loc[opportunity_cat, "avg_favorites"]:.1f}) with only {category_stats.loc[opportunity_cat, "product_count"]:.0f} products'
                })
        
        if 'trending' in locals() and len(trending) > 0:
            top_trending = trending.iloc[0]
            recommendations.append({
                'priority': 'Medium',
                'area': 'Trending Products',
                'action': f'Increase inventory for trending products',
                'reason': f'Product getting {top_trending["favorites_30d"]:,} recent favorites'
            })
        
        for rec in recommendations:
            st.markdown(f"""
            <div class="insight-box">
                <p><strong>[{rec['priority']}] {rec['area']}</strong></p>
                <p>Action: {rec['action']}</p>
                <p>Why: {rec['reason']}</p>
            </div>
            """, unsafe_allow_html=True)

# ============================================
# FOOTER
# ============================================
st.markdown(f"""
<div class="footer">
    Ozon E-Commerce Analytics Dashboard
    <br>
    Data Source: Ozon Product Data
    <br>
    Last Updated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}
</div>
""", unsafe_allow_html=True)