import React, { useEffect, useState } from 'react';
import {
  Container, Typography, TextField, Button, MenuItem, Grid, Box, InputAdornment
} from '@mui/material';
import { Formik, Form } from 'formik';
import * as Yup from 'yup';
import { DatePicker, LocalizationProvider } from '@mui/x-date-pickers';
import { AdapterDateFns } from '@mui/x-date-pickers/AdapterDateFns';

const StartScreen = () => {
  const [projects, setProjects] = useState([]);

  useEffect(() => {
    window.api.getProjects().then(setProjects);
  }, []);

  const validationSchema = Yup.object({
    projectType: Yup.string().required('Required'),
    projectCode: Yup.string().required('Required'),
    description: Yup.string().required('Required'),
    clientName: Yup.string().required('Required'),
    location: Yup.string().required('Required'),
    valueInCrores: Yup.number().positive().required('Required'),
    startDate: Yup.date().required('Required').min(new Date(), 'Must be future date'),
    endDate: Yup.date().required('Required').min(Yup.ref('startDate'), 'Must be after start date'),
    concreteQty: Yup.number().positive().required('Required'),
    fuelCost: Yup.number().positive().required('Required'),
    powerCost: Yup.number().positive().required('Required'),
    workbookPath: Yup.string().required('Required')
  });

  const handleSubmit = async (values) => {
    await window.api.createProject({
      name: values.projectCode,
      password: 'default', // You can add password modal later
      start_date: values.startDate,
      end_date: values.endDate
    });
    setProjects(await window.api.getProjects());
  };

  return (
    <LocalizationProvider dateAdapter={AdapterDateFns}>
      <Container maxWidth="md" sx={{ mt: 4, p: 3, bgcolor: '#fff', borderRadius: 2, boxShadow: 3 }}>
        <Typography variant="h5" gutterBottom align="center">
          SHAPOORJI PALLONJI _CO. LTD
        </Typography>

        <Formik
          initialValues={{
            projectType: '',
            projectCode: '',
            description: '',
            clientName: '',
            location: '',
            valueInCrores: '',
            startDate: null,
            endDate: null,
            concreteQty: '',
            fuelCost: '',
            powerCost: '',
            workbookPath: ''
          }}
          validationSchema={validationSchema}
          onSubmit={handleSubmit}
        >
          {({ values, errors, touched, handleChange, setFieldValue }) => (
            <Form>
              <Grid container spacing={2}>
  <Grid item xs={12}>
    <TextField
      select
      fullWidth
      name="projectType"
      label="Which Project"
      value={values.projectType}
      onChange={handleChange}
      error={touched.projectType && Boolean(errors.projectType)}
      helperText={touched.projectType && errors.projectType}
    >
      <MenuItem value="new">New Project</MenuItem>
      {projects.map((p) => (
        <MenuItem key={p.id} value={p.name}>{p.name}</MenuItem>
      ))}
    </TextField>
  </Grid>

  <Grid item xs={12}>
    <TextField
      fullWidth
      name="projectCode"
      label="Project Code"
      value={values.projectCode}
      onChange={handleChange}
      error={touched.projectCode && Boolean(errors.projectCode)}
      helperText={touched.projectCode && errors.projectCode}
    />
  </Grid>

  <Grid item xs={12}>
    <TextField
      fullWidth
      name="description"
      label="Description"
      value={values.description}
      onChange={handleChange}
      error={touched.description && Boolean(errors.description)}
      helperText={touched.description && errors.description}
    />
  </Grid>

  <Grid item xs={12}>
    <TextField
      fullWidth
      name="clientName"
      label="Client Name"
      value={values.clientName}
      onChange={handleChange}
      error={touched.clientName && Boolean(errors.clientName)}
      helperText={touched.clientName && errors.clientName}
    />
  </Grid>

  <Grid item xs={12}>
    <TextField
      fullWidth
      name="location"
      label="Location"
      value={values.location}
      onChange={handleChange}
      error={touched.location && Boolean(errors.location)}
      helperText={touched.location && errors.location}
    />
  </Grid>

  <Grid item xs={12}>
    <TextField
      fullWidth
      type="number"
      name="valueInCrores"
      label="Project Value (Crores)"
      value={values.valueInCrores}
      onChange={handleChange}
      error={touched.valueInCrores && Boolean(errors.valueInCrores)}
      helperText={touched.valueInCrores && errors.valueInCrores}
    />
  </Grid>

  <Grid item xs={12}>
    <DatePicker
      label="Start Date"
      value={values.startDate}
      onChange={(val) => setFieldValue('startDate', val)}
      renderInput={(params) => (
        <TextField
          {...params}
          fullWidth
          error={touched.startDate && Boolean(errors.startDate)}
          helperText={touched.startDate && errors.startDate}
        />
      )}
    />
  </Grid>

  <Grid item xs={12}>
    <DatePicker
      label="End Date"
      value={values.endDate}
      onChange={(val) => setFieldValue('endDate', val)}
      renderInput={(params) => (
        <TextField
          {...params}
          fullWidth
          error={touched.endDate && Boolean(errors.endDate)}
          helperText={touched.endDate && errors.endDate}
        />
      )}
    />
  </Grid>

  <Grid item xs={12}>
    <TextField
      fullWidth
      type="number"
      name="concreteQty"
      label="Concrete Quantity"
      value={values.concreteQty}
      onChange={handleChange}
      error={touched.concreteQty && Boolean(errors.concreteQty)}
      helperText={touched.concreteQty && errors.concreteQty}
    />
  </Grid>

  <Grid item xs={12}>
    <TextField
      fullWidth
      type="number"
      name="fuelCost"
      label="Fuel Cost per Ltr"
      InputProps={{
        startAdornment: <InputAdornment position="start">₹</InputAdornment>
      }}
      value={values.fuelCost}
      onChange={handleChange}
      error={touched.fuelCost && Boolean(errors.fuelCost)}
      helperText={touched.fuelCost && errors.fuelCost}
    />
  </Grid>

  <Grid item xs={12}>
    <TextField
      fullWidth
      type="number"
      name="powerCost"
      label="Power Cost per Unit"
      value={values.powerCost}
      onChange={handleChange}
      error={touched.powerCost && Boolean(errors.powerCost)}
      helperText={touched.powerCost && errors.powerCost}
    />
  </Grid>

  <Grid item xs={12}>
    <TextField
      fullWidth
      name="workbookPath"
      label="Workbook Path"
      value={values.workbookPath}
      onChange={handleChange}
      error={touched.workbookPath && Boolean(errors.workbookPath)}
      helperText={touched.workbookPath && errors.workbookPath}
    />
  </Grid>

  <Grid item xs={12} sx={{ textAlign: 'center', mt: 2 }}>
    <Button type="submit" variant="contained" color="primary" sx={{ mr: 2 }}>
      Save & Proceed
    </Button>
    <Button variant="outlined" color="secondary">
      Close
    </Button>
  </Grid>
</Grid>

            </Form>
          )}
        </Formik>
      </Container>
    </LocalizationProvider>
  );
};

export default StartScreen;
