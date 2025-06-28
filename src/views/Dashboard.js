/*!

=========================================================
* Paper Dashboard React - v1.3.2
=========================================================

* Product Page: https://www.creative-tim.com/product/paper-dashboard-react
* Copyright 2023 Creative Tim (https://www.creative-tim.com)

* Licensed under MIT (https://github.com/creativetimofficial/paper-dashboard-react/blob/main/LICENSE.md)

* Coded by Creative Tim

=========================================================

* The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

*/
import React, {useState, useEffect} from "react";
// react plugin used to create charts
import { Line, Pie } from "react-chartjs-2";
// reactstrap components
import {
  Card,
  CardHeader,
  CardBody,
  CardFooter,
  CardTitle,
  Row,
  Col,
} from "reactstrap";
// core components
import {
  dashboard24HoursPerformanceChart,
  dashboardEmailStatisticsChart,
  dashboardNASDAQChart,
} from "variables/charts.js";
import * as XLSX from "xlsx";

function formatRate(rate) {
  const rounded = Math.round(rate * 100) / 100;
  return rounded % 1 === 0 ? rounded.toString() : rounded.toFixed(2).replace(/\.?0+$/, '');
}

function Dashboard() {
const [schoolCount, setSchoolCount] = useState(null);
const [schoolCountGrowthRate, setSchoolCountGrowthRate] = useState(null);

useEffect(() => {
  async function fetchSchoolStats() {
    const response = await fetch('/report.xlsx');
    const buffer = await response.arrayBuffer();
    const workbook = XLSX.read(buffer, { type: 'buffer' });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const data = XLSX.utils.sheet_to_json(sheet);

    const semesters = [...new Set(data.map(row => row['학기']))].sort();
    const latest = semesters[semesters.length - 1];
    const prev = semesters.length >= 2 ? semesters[semesters.length - 2] : null;

    const filteredLatest = data.filter(row => row['학기'] === latest);
    const latestSchools = new Set(filteredLatest.map(row => row['학교']));

    setSchoolCount(latestSchools.size);

    if (prev) {
      const filteredPrev = data.filter(row => row['학기'] === prev);
      const prevSchools = new Set(filteredPrev.map(row => row['학교']));

      const increase = latestSchools.size - prevSchools.size;
      const rate = prevSchools.size > 0
        ? ((increase / prevSchools.size) * 100).toFixed(1)
        : '∞';

      setSchoolCountGrowthRate(rate);
    } else {
      setSchoolCountGrowthRate('N/A');
    }
  }

  fetchSchoolStats();
}, []);

  return (
    <>
      <div className="content">
        <Row>
          <Col lg="3" md="6" sm="6">
            <Card className="card-stats">
              <CardBody>
                <Row>
                  <Col md="4" xs="5">
                    <div className="icon-big text-center icon-warning">
                      <i className="nc-icon nc-money-coins text-success" />
                    </div>
                  </Col>
                  <Col md="8" xs="7">
                    <div className="numbers">
                      <p className="card-category">총 매출</p>
                      <CardTitle tag="p">50,000,000원</CardTitle>
                      <p />
                    </div>
                  </Col>
                </Row>
              </CardBody>
              <CardFooter>
                <hr />
                <div className="stats" style={{ textAlign: 'center' }}>
                  전 학기 대비 <strong>50%</strong> 성장
                </div>
              </CardFooter>
            </Card>
          </Col>
          <Col lg="3" md="6" sm="6">
            <Card className="card-stats">
              <CardBody>
                <Row>
                  <Col md="4" xs="5">
                    <div className="icon-big text-center icon-warning">
                      <i className="nc-icon nc-bank text-info" />
                    </div>
                  </Col>
                  <Col md="8" xs="7">
                    <div className="numbers">
                      <p className="card-category">로컬트립가이드 참여 학교 수</p>
                      <CardTitle tag="p">{schoolCount}곳</CardTitle>
                      <p />
                    </div>
                  </Col>
                </Row>
              </CardBody>
              <CardFooter>
                <hr />
                <div className="stats" style={{ textAlign: 'center' }}>
                  전 학기 대비 <strong>{formatRate(schoolCountGrowthRate)}%</strong> 증가
                </div>
              </CardFooter>
            </Card>
          </Col>
          <Col lg="3" md="6" sm="6">
            <Card className="card-stats">
              <CardBody>
                <Row>
                  <Col md="4" xs="5">
                    <div className="icon-big text-center icon-warning">
                      <i className="nc-icon nc-vector text-danger" />
                    </div>
                  </Col>
                  <Col md="8" xs="7">
                    <div className="numbers">
                      <p className="card-category">Errors</p>
                      <CardTitle tag="p">23</CardTitle>
                      <p />
                    </div>
                  </Col>
                </Row>
              </CardBody>
              <CardFooter>
                <hr />
                <div className="stats" style={{ textAlign: 'center' }}>
                  전 학기 대비 <strong>50%</strong> 증가
                </div>
              </CardFooter>
            </Card>
          </Col>
          <Col lg="3" md="6" sm="6">
            <Card className="card-stats">
              <CardBody>
                <Row>
                  <Col md="4" xs="5">
                    <div className="icon-big text-center icon-warning">
                      <i className="nc-icon nc-favourite-28 text-danger" />
                    </div>
                  </Col>
                  <Col md="8" xs="7">
                    <div className="numbers">
                      <p className="card-category">인스타 팔로워</p>
                      <CardTitle tag="p">456</CardTitle>
                      <p />
                    </div>
                  </Col>
                </Row>
              </CardBody>
              <CardFooter>
                <hr />
                <div className="stats" style={{ textAlign: 'center' }}>
                  전 학기 대비 <strong>1,200명</strong> 증가
                </div>
              </CardFooter>
            </Card>
          </Col>
        </Row>
      </div>
    </>
  );
}

export default Dashboard;
