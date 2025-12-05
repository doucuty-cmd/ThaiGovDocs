// static/script.js

// State management
const state = {
    students: [],
    teachers: []
};

// Thai months array
const thaiMonths = [
    "มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน",
    "พฤษภาคม", "มิถุนายน", "กรกฎาคม", "สิงหาคม",
    "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม"
];

// Student functions
function addStudent() {
    const studentData = {
        title: document.getElementById('student_title').value,
        firstname: document.getElementById('student_firstname').value,
        lastname: document.getElementById('student_lastname').value,
        grade: document.getElementById('student_grade').value,
        room: document.getElementById('student_room').value,
        id: Date.now()
    };

    if (!validateStudentData(studentData)) {
        alert('กรุณากรอกชื่อและนามสกุลให้ครบถ้วน');
        return;
    }

    state.students.push(studentData);
    updateStudentList();
    clearStudentForm();
}

function validateStudentData(data) {
    return data.firstname.trim() !== '' && data.lastname.trim() !== '';
}

function removeStudent(studentId) {
    state.students = state.students.filter(student => student.id !== studentId);
    updateStudentList();
}

function updateStudentList() {
    const studentList = document.getElementById('studentList');
    studentList.innerHTML = state.students.map(student => `
        <div class="student-item">
            <div class="student-info">
                <span>${student.title}${student.firstname} ${student.lastname} 
                      นักเรียนชั้นมัธยมศึกษาปีที่ ${student.grade}/${student.room}</span>
            </div>
            <button onclick="removeStudent(${student.id})" class="btn-remove">ลบ</button>
        </div>
    `).join('');
}

function clearStudentForm() {
    const fields = ['title', 'firstname', 'lastname', 'grade', 'room'];
    fields.forEach(field => {
        const element = document.getElementById(`student_${field}`);
        if (element.tagName === 'SELECT') {
            element.selectedIndex = 0;
        } else {
            element.value = '';
        }
    });
}

// Teacher functions
function addTeacher() {
    const teacherData = {
        title: document.getElementById('teacher_title').value,
        firstname: document.getElementById('teacher_firstname').value,
        lastname: document.getElementById('teacher_lastname').value,
        department: document.getElementById('teacher_department').value,
        id: Date.now()
    };

    if (!validateTeacherData(teacherData)) {
        alert('กรุณากรอกชื่อและนามสกุลให้ครบถ้วน');
        return;
    }

    state.teachers.push(teacherData);
    updateTeacherList();
    clearTeacherForm();
    updateIssuerInfo();
}

function validateTeacherData(data) {
    return data.firstname.trim() !== '' && data.lastname.trim() !== '';
}

function removeTeacher(teacherId) {
    state.teachers = state.teachers.filter(teacher => teacher.id !== teacherId);
    updateTeacherList();
}

function updateTeacherList() {
    const teacherList = document.getElementById('teacherList');
    teacherList.innerHTML = state.teachers.map(teacher => `
        <div class="student-item">
            <div class="student-info">
                <span>${teacher.title}${teacher.firstname} ${teacher.lastname} ${teacher.department}</span>
            </div>
            <button onclick="removeTeacher(${teacher.id})" class="btn-remove">ลบ</button>
        </div>
    `).join('');
}

function clearTeacherForm() {
    const fields = ['title', 'firstname', 'lastname', 'department'];
    fields.forEach(field => {
        const element = document.getElementById(`teacher_${field}`);
        if (element.tagName === 'SELECT') {
            element.selectedIndex = 0;
        } else {
            element.value = '';
        }
    });
}

function updateIssuerInfo() {
    if (state.teachers.length > 0) {
        const firstTeacher = state.teachers[0];
        document.getElementById('issuer_name').value = 
            `${firstTeacher.title}${firstTeacher.firstname} ${firstTeacher.lastname}`;
        document.getElementById('issuer_position').value = 
            `ครู${firstTeacher.department}`;
    }
}

// Document processing functions
function collectFormData() {
    return {
        department: document.getElementById('department').value,
        date: document.getElementById('date').value,
        subject: document.getElementById('subject').value,
        activity: {
            name: document.getElementById('activity_name').value,
            location: document.getElementById('location').value,
            startDate: document.querySelector('input[name="start_date"]').value,
            endDate: document.querySelector('input[name="end_date"]').value
        },
        students: state.students,
        teachers: state.teachers,
        issuer: {
            name: document.getElementById('issuer_name').value,
            position: document.getElementById('issuer_position').value
        }
    };
}

function formatThaiDate(dateStr) {
    const date = new Date(dateStr);
    return `${date.getDate()} ${thaiMonths[date.getMonth()]} พ.ศ. ${date.getFullYear() + 543}`;
}

function updatePreview() {
    const formData = collectFormData();

    // Update basic info
    document.getElementById('preview_department').textContent = formData.department;
    document.getElementById('preview_date').textContent = formatThaiDate(formData.date);
    document.getElementById('preview_subject').textContent = `${formData.subject}${formData.activity.name}`;
    // Update issuer
    document.getElementById('preview_issuer_name').textContent = 
    formData.issuer.name ? `(${formData.issuer.name})` : '';
    document.getElementById('preview_issuer_position').textContent = 
    formData.issuer.position;

    // Generate content
    let contentText = `                ด้วย${formData.department} มีความประสงค์นำนักเรียนเข้าร่วมกิจกรรม${formData.activity.name} `;
    contentText += `ณ ${formData.activity.location} `;

    // Handle activity dates
    const startDate = new Date(formData.activity.startDate);
    const endDate = new Date(formData.activity.endDate);
    
    if (startDate.getTime() === endDate.getTime()) {
        contentText += `ในวันที่ ${formatThaiDate(startDate)}`;
    } else {
        contentText += `ระหว่างวันที่ ${startDate.getDate()} - ${endDate.getDate()} `;
        contentText += `${thaiMonths[endDate.getMonth()]} พ.ศ. ${endDate.getFullYear() + 543}`;
    }

    // Add participants
    if (state.students.length > 0 || state.teachers.length > 0) {
        contentText += '\n\n                ในกิจกรรมนี้ ได้ส่งนักเรียนและผู้ควบคุม คือ';

        // Add students
        state.students.forEach((student, index) => {
            contentText += `\n                ${index + 1}. ${student.title}${student.firstname} ${student.lastname} `;
            contentText += `นักเรียนชั้นมัธยมศึกษาปีที่ ${student.grade}/${student.room}`;
        });

        // Add teachers after students
        state.teachers.forEach((teacher, index) => {
            contentText += `\n                ${state.students.length + index + 1}. ${teacher.title}${teacher.firstname} ${teacher.lastname} ${teacher.department}`;
        });
    }

    document.getElementById('preview_content').innerHTML = contentText.replace(/\n/g, '<br>');
}

// Document generation functions
function generateDocument(type = 'docx') {
    const formData = collectFormData();
    formData.generate_pdf = (type === 'pdf');

    fetch('/generate', {
        method: 'POST',
        headers: {
            'Content-Type': 'application/json',
        },
        body: JSON.stringify(formData)
    })
    .then(response => {
        if (!response.ok) {
            throw new Error('Network response was not ok');
        }
        return response.json();
    })
    .then(data => {
        if (data.docx_path || data.pdf_path) {
            window.location.href = `/download/${type}`;
        }
    })
    .catch(error => {
        console.error('Error:', error);
        alert('เกิดข้อผิดพลาดในการสร้างเอกสาร');
    });
}

// Event listeners
document.addEventListener('DOMContentLoaded', function() {
    // Add event listeners for form changes
    const forms = document.querySelectorAll('form');
    forms.forEach(form => {
        form.addEventListener('change', updatePreview);
        form.addEventListener('input', updatePreview);
    });

    // Initial preview update
    updatePreview();
});