import java.io.IOException;

import java.util.*;

import java.io.StringReader;

import org.apache.hadoop.conf.Configuration;

import org.apache.hadoop.filecache.DistributedCache;

import org.apache.hadoop.fs.Path;

import org.apache.hadoop.io.DoubleWritable;

import org.apache.hadoop.io.IntWritable;

import org.apache.hadoop.io.LongWritable;

import org.apache.hadoop.io.Text;

import org.apache.hadoop.mapreduce.Job;

import org.apache.hadoop.mapreduce.Mapper;

import org.apache.hadoop.mapreduce.Reducer;

import org.apache.hadoop.mapred.Reporter;

import org.apache.hadoop.mapred.*;

import org.apache.hadoop.mapreduce.lib.input.FileInputFormat;

import org.apache.hadoop.mapreduce.lib.output.FileOutputFormat;

import org.apache.hadoop.util.GenericOptionsParser;

import java.util.Iterator;

import java.util.List;

import javax.xml.bind.ParseConversionEvent;

import java.io.PrintWriter;

import java.io.File;

import java.io.BufferedInputStream;

import java.io.BufferedReader;

import java.io.FileInputStream;

import java.io.FileReader;

import java.io.InputStream;

import org.apache.commons.lang.StringUtils;

public class kmeans {
  
	public static String cent[][] = new String[500][18];
	public static String new_cent[][] = new String[500][18];
	public static int new_cent_count = 0;
	public static int iter_exit = 0;
	public static int iter_exit_count = 0;
	public static double attr[][];
	public static int row;
	public static int n_cent;
	public static String cluster="";
	
	public static class Mapper1
           extends Mapper<Object, Text, IntWritable, Text>{

		private String[][] gt;

		private double[] centroid;

		private Text word = new Text();
		
		public void map(Object key, Text value, Context context) throws IOException, InterruptedException {

        try {

			double min=0;

			String[] line = value.toString().split(","); // Gene Id with attributes 18 column

			String c = null;

			IntWritable in_key = null;

			String s="";

			for (int j = 0; j < n_cent; j++){

                double sum = 0;

                for (int k = 2; k < 18; k++){

					sum = sum + (Double.parseDouble(line[k])   - Double.parseDouble(cent[j][k]))*(Double.parseDouble(line[k])   - Double.parseDouble(cent[j][k]));

				}

                double sim=1 / (1 + Math.sqrt(sum));
	
				if(sim>min){
					
					c = null;
					
					min=sim;

					for (int k = 1; k < 18; k++){

						if(c==null){
				
							c =   line[k].toString();
					
						}
						else{
				
							c =   c + "," + line[k].toString();
			
						}
			   
						in_key = new IntWritable(j+1);

					}
					s=c;
		   
					s=in_key+ "," + s;
		   
					s=line[0] + "," + s+ "#";

					c =   line[0] + "," + c;

				}

			}
			
			cluster=cluster + s;
		
			word.set(c);
			
			context.write(in_key, word);// centroid id with attr and point with id and attr

		} catch (Exception e) {}

	}

	}

 public static class Reducer1 extends Reducer<IntWritable, Text, IntWritable, Text> {

	 private Text word = new Text();

	@SuppressWarnings("null")

	public void reduce(IntWritable key, Iterable<Text> values, Context context

				) throws IOException, InterruptedException {

		IntWritable k= new IntWritable();     

        double[] sum = new double[18];

		String c = null;

        int count=0;

        int length = 18;

        for (Text val : values){

			String[] line = val.toString().split(",");
		
        	for(int j = 2; j<line.length;j++){

        		sum[j]= Double.parseDouble(line[j]) + sum[j];                	

        	}

        	count++;	

        }
		
		int kn=key.get();
		
		for(int j = 2; j<18;j++){

            sum[j]= Double.parseDouble(cent[kn-1][j]) + sum[j];                	

        }	
		count++;	

        for(int j = 2; j< length;j++){

        	sum[j]=sum[j]/count;
			
			if(c==null){
				
				c =   String.valueOf(sum[j]);
			
			}
			
			else{
        	
				c =   c + "," + String.valueOf(sum[j]);		
		
			}

        }

		c =   "0" + "," + c;
	
		c =   key.toString() + "," + c;
	
		new_cent[new_cent_count][0] = key.toString();
	
		new_cent[new_cent_count][1] = "0";
        
		String cnt=cent[new_cent_count][0];
	
		String[] test = c.split(",");
	
		for(int i = 2; i< 18;i++){
	
			new_cent[new_cent_count][i] = test[i];		
		
		}
	
		if(new_cent_count == n_cent-1){
	
			int test_count = 0;
	
			for(int y=0;y<n_cent;y++){
				
				for(int i = 2; i< 18;i++){
			
					if(Double.parseDouble(cent[y][i]) == Double.parseDouble(new_cent[y][i])){
				
						test_count++;
			
					}
			
				}
	
			}
			if(test_count == (n_cent*16)-1){
		
				iter_exit=1;
	
			}
	
			if(iter_exit_count == n_cent-1){
			
			iter_exit =1;
			
		}
		else{
		
			iter_exit_count = 0;
		
			new_cent_count = 0;
                        
			for(int i =0;i<n_cent;i++){
				
				cent[i][0] = Integer.toString(i+1);
			
				cent[i][1] = cent[i][0];

				for(int j =2;j<18;j++)	{
					
					cent[i][j] = new_cent[i][j];
					
				}
		
			}			
			
		}
	}
	else{
	 
		new_cent_count++;  

	}
	
	if(kn==n_cent){
		
		iter_exit_count = 0;
	
	}
	
	word.set(c);

    context.write(key, word);

        }

 }





  public static void main(String[] args) throws Exception {

    

    

    int col=0; 
    int iter = 0;
    row=0;
	
    n_cent = Integer.parseInt(args[2]);

    String file = args[1];

      String temp1="p1";

    String temp2="prb1.2";
FileReader filereader1 = new FileReader("/home/hadoop/hadoop/Centroid.csv");
		BufferedReader in1 = new BufferedReader(filereader1);
		int count = 0;
		String line = null;
		
        while((line = in1.readLine()) != null && count<n_cent) {
        	String[] query1 = line.split(",");
        	String query2  = null;
        	
		cent[count][0]=query1[0];
		cent[count][1]="0";
        	for(int j=2; j<18;j++)
        	{
            		//cent[count][j]=Double.parseDouble(query1[j]);
			cent[count][j]=query1[j];
        	}
    		count++;
        }
			Configuration conf = new Configuration();

        		Job job = Job.getInstance(conf, "kmeans_clustering");

        		job.setJarByClass(kmeans.class);

        		job.setMapperClass(Mapper1.class);

        		//job.setCombinerClass(Reducer1.class);

        		job.setReducerClass(Reducer1.class);

        		job.setOutputKeyClass(IntWritable.class);

        		job.setOutputValueClass(Text.class);

        		FileInputFormat.addInputPath(job, new Path(args[0]));

        		FileOutputFormat.setOutputPath(job, new Path(args[1]));

        		job.waitForCompletion(true);

   			//System.exit(job.waitForCompletion(true) ? 0 : 1);
	
	
	while(iter_exit!=1)
	{
			cluster="";
			file = file + Integer.toString(iter);
			Configuration conf1 = new Configuration();

        		Job job1 = Job.getInstance(conf1, "kmeans_clustering");

        		job1.setJarByClass(kmeans.class);

        		job1.setMapperClass(Mapper1.class);

        		//job1.setCombinerClass(Reducer1.class);

        		job1.setReducerClass(Reducer1.class);

        		job1.setOutputKeyClass(IntWritable.class);

        		job1.setOutputValueClass(Text.class);

        		FileInputFormat.addInputPath(job1, new Path(args[0]));

        		FileOutputFormat.setOutputPath(job1, new Path(file));

        		job1.waitForCompletion(true);
			iter++;
			
   			//System.exit(job1.waitForCompletion(true) ? 0 : 1);
			
		
	} 
        PrintWriter pw = new PrintWriter(new File("/home/hadoop/hadoop/test3.csv"));
        String[] l = cluster.split("#");
	int gt[][]=new int[l.length+1][3];
	int g_cl[][]=new int[l.length+1][3];
	for(int i=0;i<l.length;i++)
	{
		StringBuilder sb = new StringBuilder();
		sb.append(l[i]);
		sb.append('\n');
		pw.write(sb.toString());
		String[] l1 = l[i].split(",");
		gt[i+1][1]=Integer.valueOf(l1[0]);
  		gt[i+1][2]=Integer.valueOf(l1[2]);
		g_cl[i+1][1]=Integer.valueOf(l1[0]);
		g_cl[i+1][2]=Integer.valueOf(l1[1]);
		
	}
                
        pw.close();
    
	int[][] gd_matx= new int[l.length+1][l.length+1];
            int[][] cl_matx = new int[l.length+1][l.length+1];
            double m11 = 0, m00 = 0, m01 = 0, m10 = 0;

            for (int i = 1; i < l.length + 1; i++)
            {
                for (int j = 1; j < l.length + 1; j++)
                {
                    if (gt[i][2] == gt[j][2])
                    {
                        gd_matx[i][j] = 1;
                    }
                    else
                    {
                        gd_matx[i][j] = 0;
                    }
                    if (g_cl[i - 1][2] == g_cl[j - 1][2])
                    {
                        cl_matx[i][j] = 1;
                        if (gd_matx[i][j] == 1)
                        {

                            m11++;
                        }
                        else if (gd_matx[i][j] == 0)
                        {
                            m10++;
                        }
                    }
                    else
                    {
                        cl_matx[i][j] = 0;
                        if (gd_matx[i][j] == 1)
                        {
                            m01++;
                        }
                        else if (gd_matx[i][j] == 0)
                        {
                            m00++;
                        }
                    }

                }
            }

            double rand;
            rand = ((m11 + m00) / (m11 + m00 + m01 + m10));
            double jaccof;
            jaccof =((m11) / (m11 + m01 + m10));	

System.out.println("Jaccard Coeff: " + jaccof);
System.out.println("Rand Index: " + rand);
		
	System.exit(1);

  }

}