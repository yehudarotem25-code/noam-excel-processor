import streamlit as st
import pandas as pd
from io import BytesIO

# הגדרת עיצוב הדף
st.set_page_config(
    page_title="מעבד קבצי Excel",
    page_icon="📊",
    layout="wide"
)

# כותרת ראשית
st.title("📊 מעבד קבצי Excel - חיבור וצבירת נתונים")
st.markdown("---")

# הסבר למשתמש
st.markdown("""
### הוראות שימוש:
1. העלה מספר קבצי Excel מסוג A (Extensions ו-Endpoints) - המכילים עמודת **QuestionnaireFillID**
2. העלה קובץ Excel אחד מסוג B (קובץ ההתייחסות) - המכיל **מספר תשובון**, **סטטוס תשובון**, **תאריך יצירה**
3. לחץ על כפתור "עבד נתונים"
4. הורד את הקובץ המעובד
""")

st.markdown("---")

# פונקציה להמרת DataFrame ל-Excel להורדה
def convert_df_to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        # גליון 1: תוצאות מלאות
        df.to_excel(writer, index=False, sheet_name='תוצאות מלאות')
        
        # גליון 2: סיכום לפי סטטוס
        summary_df = df.groupby('סטטוס תשובון').agg({
            'מספר תשובון': 'count',  # כמה IDs בכל סטטוס
            'Count': 'sum'  # סה"כ מופעים בכל סטטוס
        }).reset_index()
        
        # שינוי שמות העמודות לברורים יותר
        summary_df.columns = ['סטטוס תשובון', 'כמות תשובונים (IDs)', 'סה"כ מופעים']
        
        # הוספת שורת סיכום
        total_row = pd.DataFrame({
            'סטטוס תשובון': ['סה"כ'],
            'כמות תשובונים (IDs)': [summary_df['כמות תשובונים (IDs)'].sum()],
            'סה"כ מופעים': [summary_df['סה"כ מופעים'].sum()]
        })
        summary_df = pd.concat([summary_df, total_row], ignore_index=True)
        
        summary_df.to_excel(writer, index=False, sheet_name='סיכום לפי סטטוס')
        
    output.seek(0)
    return output

# פונקציה לטעינת קבצי Type A
def load_type_a_files(uploaded_files):
    """טוען ומאחד מספר קבצי Excel"""
    all_data = []
    
    for file in uploaded_files:
        try:
            # קריאת הקובץ
            df = pd.read_excel(file)
            
            # וידוא שקיימת עמודת QuestionnaireFillID
            if 'QuestionnaireFillID' not in df.columns:
                st.warning(f"⚠️ הקובץ {file.name} לא מכיל עמודת 'QuestionnaireFillID' - מדלג")
                continue
            
            # המרת QuestionnaireFillID למחרוזת (למניעת בעיות עם אפסים מובילים)
            df['QuestionnaireFillID'] = df['QuestionnaireFillID'].astype(str)
            
            all_data.append(df)
            st.success(f"✅ הקובץ {file.name} נטען בהצלחה ({len(df)} שורות)")
            
        except Exception as e:
            st.error(f"❌ שגיאה בטעינת הקובץ {file.name}: {str(e)}")
    
    if all_data:
        # איחוד כל הקבצים
        combined_df = pd.concat(all_data, ignore_index=True)
        return combined_df
    else:
        return None

# פונקציה לטעינת קובץ Type B
def load_type_b_file(uploaded_file):
    """טוען את קובץ ההתייחסות"""
    try:
        df = pd.read_excel(uploaded_file)
        
        # בדיקת עמודות חובה
        required_columns = ['מספר תשובון', 'סטטוס תשובון', 'תאריך יצירה']
        missing_columns = [col for col in required_columns if col not in df.columns]
        
        if missing_columns:
            st.error(f"❌ הקובץ חסר עמודות: {', '.join(missing_columns)}")
            return None
        
        # המרת מספר תשובון למחרוזת
        df['מספר תשובון'] = df['מספר תשובון'].astype(str)
        
        st.success(f"✅ קובץ ההתייחסות נטען בהצלחה ({len(df)} שורות)")
        return df
        
    except Exception as e:
        st.error(f"❌ שגיאה בטעינת קובץ ההתייחסות: {str(e)}")
        return None

# פונקציה לעיבוד הנתונים
def process_data(type_a_df, type_b_df):
    """מבצע את לוגיקת העיבוד: ספירה, איחוד ויצירת הטבלה הסופית"""
    
    # שלב 1: ספירת מופעים של כל QuestionnaireFillID
    st.info("📊 סופר מופעים של כל QuestionnaireFillID...")
    count_df = type_a_df.groupby('QuestionnaireFillID').size().reset_index(name='Count')
    
    st.success(f"✅ נמצאו {len(count_df)} QuestionnaireFillIDs ייחודיים")
    
    # שלב 2: Inner Join עם קובץ B (רק תשובונים שמופיעים בשני הקבצים)
    st.info("🔗 מבצע איחוד עם קובץ ההתייחסות (רק תשובונים משותפים)...")
    result_df = type_b_df.merge(
        count_df, 
        left_on='מספר תשובון', 
        right_on='QuestionnaireFillID', 
        how='inner'  # Inner join - רק תשובונים שמופיעים גם בקבצי A וגם בקובץ B
    )
    
    # מחיקת עמודת QuestionnaireFillID הכפולה (נשאר רק מספר תשובון)
    if 'QuestionnaireFillID' in result_df.columns:
        result_df = result_df.drop('QuestionnaireFillID', axis=1)
    
    # שלב 3: סידור העמודות לפי הדרישה: מספר תשובון, תאריך יצירה, סטטוס תשובון, Count
    result_df = result_df[['מספר תשובון', 'תאריך יצירה', 'סטטוס תשובון', 'Count']]
    
    # שלב 4: עיצוב עמודת התאריך
    st.info("📅 מעצב תאריכים...")
    # המרת התאריך לפורמט DD/MM/YYYY
    result_df['תאריך יצירה'] = pd.to_datetime(result_df['תאריך יצירה']).dt.strftime('%d/%m/%Y')
    
    st.success(f"✅ עיבוד הושלם! הטבלה הסופית מכילה {len(result_df)} שורות")
    
    return result_df

# ממשק המשתמש
col1, col2 = st.columns(2)

with col1:
    st.subheader("📁 קבצי Type A")
    st.caption("קבצים המכילים QuestionnaireFillID")
    type_a_files = st.file_uploader(
        "העלה קבצי Extensions ו-Endpoints (ניתן להעלות מספר קבצים)",
        type=['xlsx', 'xls'],
        accept_multiple_files=True,
        key="type_a"
    )

with col2:
    st.subheader("📄 קובץ Type B")
    st.caption("קובץ התייחסות עם מספר תשובון, סטטוס, תאריך")
    type_b_file = st.file_uploader(
        "העלה קובץ התייחסות",
        type=['xlsx', 'xls'],
        accept_multiple_files=False,
        key="type_b"
    )

st.markdown("---")

# כפתור עיבוד
if st.button("🚀 עבד נתונים", type="primary", use_container_width=True):
    
    # בדיקת תקינות
    if not type_a_files:
        st.error("❌ יש להעלות לפחות קובץ Type A אחד")
    elif not type_b_file:
        st.error("❌ יש להעלות קובץ Type B")
    else:
        with st.spinner("מעבד נתונים..."):
            # טעינת הקבצים
            st.subheader("📥 טוען קבצים...")
            type_a_df = load_type_a_files(type_a_files)
            type_b_df = load_type_b_file(type_b_file)
            
            if type_a_df is not None and type_b_df is not None:
                st.markdown("---")
                
                # עיבוד הנתונים
                st.subheader("⚙️ מעבד נתונים...")
                result_df = process_data(type_a_df, type_b_df)
                
                st.markdown("---")
                
                # הצגת התוצאות
                st.subheader("📋 תוצאות:")
                
                # סטטיסטיקות
                col1, col2, col3 = st.columns(3)
                with col1:
                    st.metric("סה\"כ מספרי תשובון", len(result_df))
                with col2:
                    st.metric("תשובונים עם נתונים", len(result_df[result_df['Count'] > 0]))
                with col3:
                    st.metric("סה\"כ מופעים", result_df['Count'].sum())
                
                # הצגת הטבלה המלאה
                st.dataframe(result_df, use_container_width=True, height=400)
                
                st.markdown("---")
                
                # הצגת סיכום לפי סטטוס
                st.subheader("📊 סיכום לפי סטטוס:")
                
                # יצירת טבלת סיכום
                summary_by_status = result_df.groupby('סטטוס תשובון').agg({
                    'מספר תשובון': 'count',
                    'Count': 'sum'
                }).reset_index()
                summary_by_status.columns = ['סטטוס תשובון', 'כמות תשובונים (IDs)', 'סה"כ מופעים']
                
                # הצגת הסיכום
                st.dataframe(summary_by_status, use_container_width=True, hide_index=True)
                
                # כפתור הורדה
                excel_file = convert_df_to_excel(result_df)
                st.download_button(
                    label="📥 הורד קובץ Excel",
                    data=excel_file,
                    file_name="processed_results.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary",
                    use_container_width=True
                )
                
                st.balloons()

# פוטר
st.markdown("---")
st.markdown(
    """
    <div style='text-align: center; color: gray;'>
    <small>מערכת לעיבוד וחיבור קבצי Excel | נבנה עם Streamlit</small>
    </div>
    """,
    unsafe_allow_html=True
)
