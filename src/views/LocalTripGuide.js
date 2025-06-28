import React, { useState, useEffect } from "react";
import { Line } from "react-chartjs-2";
import {
  Card,
  CardHeader,
  CardBody,
  CardTitle,
  Row,
  Col,
  Table,
} from "reactstrap";
import * as XLSX from "xlsx";
import {
  Chart as ChartJS,
  CategoryScale,
  LinearScale,
  BarElement,
  PointElement,
  LineElement,
  Tooltip,
  Legend,
} from "chart.js";

ChartJS.register(
  CategoryScale,
  LinearScale,
  BarElement,
  PointElement,
  LineElement,
  Tooltip,
  Legend
);

function LocalTripGuide() {
  const [schoolData, setSchoolData] = useState(null);
  const [participantData, setParticipantData] = useState(null);
  const [options, setOptions] = useState(null);

  const [schools, setSchools] = useState([]);
  const [semesters, setSemesters] = useState([]);
  const [participationGrid, setParticipationGrid] = useState([]);

  useEffect(() => {
    async function fetchChartData() {
      const response = await fetch("/report.xlsx");
      const buffer = await response.arrayBuffer();
      const workbook = XLSX.read(buffer, { type: "buffer" });
      const sheet = workbook.Sheets[workbook.SheetNames[0]];
      const rawData = XLSX.utils.sheet_to_json(sheet);

      const semesters = [...new Set(rawData.map((row) => row["학기"]))].sort();
      const schools = [...new Set(rawData.map((row) => row["학교"]))].sort();

      // 학기별 참여 학교 수
      const schoolCountsPerSemester = semesters.map((semester) => {
        const schoolsInSemester = rawData
          .filter((row) => row["학기"] === semester)
          .map((row) => row["학교"]);
        return new Set(schoolsInSemester).size;
      });

      // 학기별 참여자 수
      const participantCountsPerSemester = semesters.map((semester) => {
        const filtered = rawData.filter((row) => row["학기"] === semester);
        const uniqueParticipants = new Set();
        filtered.forEach((row) => {
          const name =
            (row["성명(한글)"] && row["성명(한글)"].trim()) ||
            (row["성명(영문)"] && row["성명(영문)"].trim());
          if (name) uniqueParticipants.add(name);
        });
        return uniqueParticipants.size;
      });

      // 참여 여부 그리드 생성 (세로: 학기, 가로: 학교)
      const grid = semesters.map((semester) =>
        schools.map((school) =>
          rawData.some((row) => row["학기"] === semester && row["학교"] === school)
        )
      );

      setSchoolData({
        labels: semesters,
        datasets: [
          {
            label: "참여 학교 수",
            data: schoolCountsPerSemester,
            fill: false,
            borderColor: "#51CACF",
            backgroundColor: "transparent",
            pointBorderColor: "#51CACF",
            pointRadius: 4,
            pointHoverRadius: 4,
            pointBorderWidth: 8,
            tension: 0,
          },
        ],
      });

      setParticipantData({
        labels: semesters,
        datasets: [
          {
            label: "참여자 수",
            data: participantCountsPerSemester,
            fill: false,
            borderColor: "#fbc658",
            backgroundColor: "transparent",
            pointBorderColor: "#fbc658",
            pointRadius: 4,
            pointHoverRadius: 4,
            pointBorderWidth: 8,
            tension: 0,
          },
        ],
      });

      setOptions({
        plugins: { legend: { display: false } },
        scales: {
          y: {
            beginAtZero: true,
            precision: 0,
            title: { display: false, text: "수" },
          },
          x: {
            title: { display: false, text: "학기" },
          },
        },
        responsive: true,
        maintainAspectRatio: false,
      });

      setSchools(schools);
      setSemesters(semesters);
      setParticipationGrid(grid);
    }

    fetchChartData();
  }, []);

  if (!schoolData || !participantData || !options || !participationGrid.length) {
    return <div>Loading chart...</div>;
  }

  return (
    <div className="content">
      <h3 style={{ margin: '12px 0' }}>학기별 통계</h3>
      <Row>
        <Col md="6">
          <Card className="card-chart">
            <CardHeader>
              <div style={{ display: 'flex', flexDirection: 'row', alignItems: 'flex-end', gap: '10px' }}>
                <CardTitle tag="h5">학기별 참여 학교 수 추이</CardTitle>
                <p style={{ marginBottom: '1px' }} className="card-category">총 {schoolData.labels.length}학기</p>
              </div>
            </CardHeader>
            <CardBody>
              <Line data={schoolData} options={options} />
            </CardBody>
          </Card>
        </Col>
        <Col md="6">
          <Card className="card-chart">
            <CardHeader>
              <div style={{ display: 'flex', flexDirection: 'row', alignItems: 'flex-end', gap: '10px' }}>
                <CardTitle tag="h5">학기별 참여자 수 추이</CardTitle>
                <p style={{ marginBottom: '1px' }} className="card-category">총 {participantData.labels.length}학기</p>
              </div>
            </CardHeader>
            <CardBody>
              <Line data={participantData} options={options} />
            </CardBody>
          </Card>
        </Col>
      </Row>

      <Row>
        <Col md="12" style={{ overflowX: "auto", maxHeight: "600px" }}>
          <Card>
            <CardHeader>
              <div style={{ display: 'flex', flexDirection: 'row', alignItems: 'flex-end', gap: '10px' }}>
                <CardTitle tag="h5">학기별 이용 지속 히트맵</CardTitle>
                <p style={{ marginBottom: '13px' }} className="card-category">파란색: 참여, 빈칸: 미참여</p>
              </div>
            </CardHeader>
            <CardBody>
              <Table
                bordered
                responsive
                size="sm"
                style={{ tableLayout: "fixed", minWidth: schools.length * 80 }}
              >
                <thead>
                  <tr>
                    <th style={{ fontWeight: '500', textAlign: 'center', minWidth: 100 }}>학기 / 학교</th>
                    {schools.map((school) => (
                      <th
                        key={school}
                        style={{
                          minWidth: 80,
                          fontWeight: '500',
                          textAlign: "center",
                          whiteSpace: "nowrap",
                        }}
                      >
                        {school}
                      </th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {semesters.map((semester, rowIndex) => (
                    <tr key={semester}>
                      <td
                        style={{ textAlign: "center", whiteSpace: "nowrap" }}
                      >
                        {semester}
                      </td>
                      {schools.map((school, colIndex) => (
                        <td
                          key={school}
                          style={{
                            backgroundColor: participationGrid[rowIndex][colIndex]
                              ? "#51CACF"
                              : "transparent",
                            color: participationGrid[rowIndex][colIndex]
                              ? "white"
                              : "black",
                            textAlign: "center",
                            userSelect: "none",
                          }}
                        >
                          {participationGrid[rowIndex][colIndex] ? "" : ""}
                        </td>
                      ))}
                    </tr>
                  ))}
                </tbody>
              </Table>
            </CardBody>
          </Card>
        </Col>
      </Row>
    </div>
  );
}

export default LocalTripGuide;
