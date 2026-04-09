from __future__ import annotations

from typing import Any

from googleapiclient.discovery import build


class ClassroomClient:
    def __init__(self, credentials):
        self.service = build("classroom", "v1", credentials=credentials)

    def list_courses(self) -> list[dict[str, Any]]:
        """
        Recupera cursos visibles para la cuenta autenticada.
        """
        response = self.service.courses().list(pageSize=50).execute()
        return response.get("courses", [])

    def list_coursework(self, course_id: str) -> list[dict[str, Any]]:
        """
        Recupera actividades de un curso.
        """
        response = (
            self.service.courses()
            .courseWork()
            .list(courseId=course_id, pageSize=100)
            .execute()
        )
        return response.get("courseWork", [])

    def list_student_submissions(
        self,
        course_id: str,
        coursework_id: str,
    ) -> list[dict[str, Any]]:
        """
        Recupera entregas de estudiantes para una actividad.
        """
        response = (
            self.service.courses()
            .courseWork()
            .studentSubmissions()
            .list(courseId=course_id, courseWorkId=coursework_id, pageSize=100)
            .execute()
        )
        return response.get("studentSubmissions", [])
