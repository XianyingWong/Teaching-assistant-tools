# 开发时间2021/11/16 15:56
from selenium import webdriver

# 创建 WebDriver 对象，指明使用chrome浏览器驱动
wd = webdriver.Chrome(r'D:\driver\chromedriver.exe')

# 调用WebDriver 对象的get方法 可以让浏览器打开指定网址
wd.get('https://elearning.szu.edu.cn/webapps/assignment/gradeAssignmentRedirector?outcomeDefinitionId=_159741_1&currentAttemptIndex=1&numAttempts=116&anonymousMode=false&sequenceId=_43623_1_0&course_id=_43623_1&source=cp_gradebook2_view_grade_details&viewInfo=%E5%AE%8C%E6%95%B4%E7%9A%84%E6%88%90%E7%BB%A9%E4%B8%AD%E5%BF%83&attempt_id=_1494885_1&courseMembershipId=_2001986_1&cancelGradeUrl=%2Fwebapps%2Fgradebook%2Fdo%2Finstructor%2FviewGradeDetails%3Fcourse_id%3D_43623_1%26focus_cell_id%3Dcell_1_7%26courseMembershipId%3D_2001986_1%26outcomeDefinitionId%3D_159741_1%26attemptId%3D_1494885_1&submitGradeUrl=%2Fwebapps%2Fgradebook%2Fdo%2Finstructor%2FperformGrading%3Fcourse_id%3D_43623_1%26cmd%3Dnext%26sequenceId%3D_43623_1_0')
